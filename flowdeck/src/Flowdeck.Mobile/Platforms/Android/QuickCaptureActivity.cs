using Android.App;
using Android.Content.PM;
using Android.OS;
using Android.Text;
using Android.Views;
using Android.Views.InputMethods;
using Android.Widget;
using AndroidX.AppCompat.App;
using Flowdeck.Core.Models;
using Flowdeck.Core.Parsing;
using Flowdeck.Mobile.Services;
using Flowdeck.Mobile.Widgets;

// This is a plain Android screen, not a MAUI one, and MAUI has a Button of its own in the
// implicit usings. Named outright so the ambiguity cannot come back.
using Button = Android.Widget.Button;

namespace Flowdeck.Mobile;

/// <summary>
/// The capture popup the home-screen strip opens: a box, a running reading of it, and save.
///
/// A plain Android activity rather than a MAUI page, because it has to be in front of the
/// user in the time it takes to have the thought. Starting the whole app first would spend
/// that time on a splash screen, and the idea this exists to catch would be gone.
/// </summary>
[Activity(
    Label = "Flowdeck",
    Theme = "@style/Flowdeck.QuickCapture",
    Exported = true,
    LaunchMode = LaunchMode.SingleTop,
    ExcludeFromRecents = true,
    ConfigurationChanges = ConfigChanges.ScreenSize | ConfigChanges.Orientation | ConfigChanges.UiMode)]

// Puts Flowdeck in the popup that comes up over selected text, next to copy and share. The
// label above is what that entry reads, and the menu is narrow, so it is the app's name and
// nothing else.
[IntentFilter(
    new[] { "android.intent.action.PROCESS_TEXT" },
    Categories = new[] { Android.Content.Intent.CategoryDefault },
    DataMimeType = "text/plain")]

// And in the share sheet, which is the way back in from an app whose selection popup does
// not show us — either because it draws its own, or because the menu ran out of room.
[IntentFilter(
    new[] { Android.Content.Intent.ActionSend },
    Categories = new[] { Android.Content.Intent.CategoryDefault },
    DataMimeType = "text/plain")]
public sealed class QuickCaptureActivity : AppCompatActivity
{
    private EditText? _input;
    private TextView? _preview;
    private Button? _save;

    private ParsedEntry _parsed = new();
    private EntryTarget? _forced;
    private bool _ready;

    protected override void OnCreate(Bundle? savedInstanceState)
    {
        base.OnCreate(savedInstanceState);

        SetContentView(Resource.Layout.quick_capture);

        _input = FindViewById<EditText>(Resource.Id.capture_input);
        _preview = FindViewById<TextView>(Resource.Id.capture_preview);
        _save = FindViewById<Button>(Resource.Id.capture_save);

        _save!.Click += (_, _) => Commit();
        FindViewById<Button>(Resource.Id.capture_cancel)!.Click += (_, _) => Finish();

        FindViewById<Button>(Resource.Id.capture_todo)!.Click += (_, _) => Force(EntryTarget.Todo);
        FindViewById<Button>(Resource.Id.capture_calendar)!.Click += (_, _) => Force(EntryTarget.Calendar);
        FindViewById<Button>(Resource.Id.capture_both)!.Click += (_, _) => Force(EntryTarget.Both);

        _input!.AfterTextChanged += (_, _) => Reparse();

        // Enter saves, the way it does in the desktop overlay.
        _input.EditorAction += (_, e) =>
        {
            if (e.ActionId is ImeAction.Done or ImeAction.Go or ImeAction.Send)
            {
                Commit();
                e.Handled = true;
            }
        };

        _save.Enabled = false;
        Fill(Intent);
        Prepare();
    }

    /// <summary>
    /// A second selection sent over while the popup is already up. SingleTop means the
    /// activity is reused rather than stacked, so without this the new text never arrives.
    /// </summary>
    protected override void OnNewIntent(Android.Content.Intent? intent)
    {
        base.OnNewIntent(intent);

        Intent = intent;
        Fill(intent);
        Reparse();
    }

    /// <summary>
    /// Puts text handed over by another app into the box, ready to be edited before it is
    /// saved. Selected text is rarely a finished todo — a date or a !TD is usually wanted —
    /// and saving it outright would take away the one moment to add them.
    ///
    /// Folded to one line by the same call the parser makes, so the box shows what will be
    /// read from it. The cursor goes to the end rather than over the text: the next keystroke
    /// should add to what arrived, not replace it.
    /// </summary>
    private void Fill(Android.Content.Intent? intent)
    {
        if (intent is null || _input is null) return;

        var text = intent.Action switch
        {
            "android.intent.action.PROCESS_TEXT" =>
                intent.GetStringExtra("android.intent.extra.PROCESS_TEXT"),
            Android.Content.Intent.ActionSend =>
                intent.GetStringExtra(Android.Content.Intent.ExtraText),
            _ => null,
        };

        if (string.IsNullOrWhiteSpace(text)) return;

        _input.Text = NaturalLanguageParser.OneLine(text);
        _input.SetSelection(_input.Text!.Length);
    }

    /// <summary>
    /// The workspace has to be on disk before anything can be saved, but the box should be
    /// typeable immediately — so it is, and only saving waits.
    /// </summary>
    private async void Prepare()
    {
        try
        {
            await Workspace.EnsureLoadedAsync();
            _ready = true;
            Reparse();

            _input?.RequestFocus();
            (GetSystemService(InputMethodService) as InputMethodManager)
                ?.ShowSoftInput(_input, ShowFlags.Implicit);
        }
        catch (Exception)
        {
            if (_preview is not null) _preview.Text = "저장소를 열지 못했습니다";
        }
    }

    private void Force(EntryTarget target)
    {
        _forced = _forced == target ? null : target;
        Reparse();
    }

    private void Reparse()
    {
        if (!_ready || _input is null || _preview is null || _save is null) return;

        var text = _input.Text ?? string.Empty;
        _parsed = Workspace.Parser.Parse(text, DateTime.Now);

        var hasInput = !string.IsNullOrWhiteSpace(text);
        _save.Enabled = hasInput;

        if (!hasInput)
        {
            _preview.Text = string.Empty;
            return;
        }

        var entry = Effective();
        var parts = new List<string> { entry.DescribeTarget(), entry.DescribeSchedule() };
        if (entry.Recurrence.IsRepeating) parts.Add(entry.Recurrence.Describe());
        if (entry.Tags.Count > 0) parts.Add(string.Join(" ", entry.Tags.Select(t => "#" + t)));

        _preview.Text = entry.Title + "\n" + string.Join(" · ", parts);
    }

    private ParsedEntry Effective()
    {
        if (_forced is null || _parsed.IsEmpty) return _parsed;

        return new ParsedEntry
        {
            RawInput = _parsed.RawInput,
            Title = _parsed.Title,
            Target = _forced.Value,
            TargetWasExplicit = true,
            Start = _parsed.Start,
            End = _parsed.End,
            HasTime = _parsed.HasTime,
            Priority = _parsed.Priority,
            Tags = _parsed.Tags,
            Recurrence = _parsed.Recurrence,
            ReminderMinutesBefore = _parsed.ReminderMinutesBefore,
            OpenTeamsMeeting = _parsed.OpenTeamsMeeting,
            Attendees = _parsed.Attendees,
        };
    }

    private async void Commit()
    {
        if (!_ready || _save is null || !_save.Enabled) return;

        _save.Enabled = false;

        try
        {
            var entry = Effective();
            var result = await Workspace.Repository.CaptureAsync(entry, DateTime.Now);
            if (result.IsEmpty)
            {
                _save.Enabled = true;
                return;
            }

            Toast.MakeText(this, "저장했습니다 · " + entry.Title, ToastLength.Short)?.Show();
            FlowdeckWidgets.RefreshAll(this);
            Finish();
        }
        catch (Exception)
        {
            if (_preview is not null) _preview.Text = "저장하지 못했습니다";
            _save.Enabled = true;
        }
    }

    /// <summary>Tapping outside the card dismisses it, as a popup should.</summary>
    public override bool OnTouchEvent(MotionEvent? e)
    {
        if (e?.Action == MotionEventActions.Down)
        {
            Finish();
            return true;
        }

        return base.OnTouchEvent(e);
    }
}
