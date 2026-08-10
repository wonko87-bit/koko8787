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

namespace Flowdeck.Mobile;

/// <summary>
/// The capture popup the home-screen strip opens: a box, a running reading of it, and save.
///
/// A plain Android activity rather than a MAUI page, because it has to be in front of the
/// user in the time it takes to have the thought. Starting the whole app first would spend
/// that time on a splash screen, and the idea this exists to catch would be gone.
/// </summary>
[Activity(
    Theme = "@style/Flowdeck.QuickCapture",
    Exported = true,
    LaunchMode = LaunchMode.SingleTop,
    ExcludeFromRecents = true,
    ConfigurationChanges = ConfigChanges.ScreenSize | ConfigChanges.Orientation | ConfigChanges.UiMode)]
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
        Prepare();
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
