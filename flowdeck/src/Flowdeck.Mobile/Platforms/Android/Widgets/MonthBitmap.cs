using Android.Graphics;
using Flowdeck.Mobile.Services;

// MAUI puts its own drawing types in the implicit usings, and several of them share a name
// with the Android ones. Spelled out here so the file cannot mean the wrong one.
using Color = Android.Graphics.Color;
using Paint = Android.Graphics.Paint;

namespace Flowdeck.Mobile.Widgets;

/// <summary>
/// Draws the month onto a bitmap for the home-screen widget.
///
/// A widget is built from RemoteViews, which only understands a short list of layouts and
/// cannot be told to lay out forty-two cells the way the app does. Drawing the grid once
/// and handing over the picture sidesteps that entirely: the widget is a single image, and
/// what it looks like is decided here, in one place, with the same rules the app uses.
///
/// Drawn at the size the widget actually is, in pixels, and in its actual shape. A fixed
/// square picture was fine on a phone whose launcher cells happen to be square, and wrong
/// everywhere else: on a tablet the cells are wider than tall, so the square sat in the
/// middle of the widget with bands of nothing above and below it, and at fewer pixels than
/// the widget had, so it was scaled up soft. Every measure below is taken from the cell
/// size, so the grid fills whatever it is given and the text stays in proportion to it.
/// </summary>
internal static class MonthBitmap
{
    private const int Columns = 7;
    private const int Rows = 6;

    /// <summary>
    /// The largest side a picture is drawn at. RemoteViews carries bitmaps across processes
    /// and refuses ones past a share of the screen's own pixel count; a tablet widget can
    /// ask for more pixels than it can show a difference for anyway.
    /// </summary>
    public const int MaxSide = 1600;

    public static Bitmap Render(int width, int height, DateTime today, bool dark)
    {
        width = Math.Clamp(width, 200, MaxSide);
        height = Math.Clamp(height, 200, MaxSide);

        var background = dark ? Color.Argb(235, 15, 15, 15) : Color.Argb(240, 255, 255, 255);
        var primary = dark ? Color.Argb(255, 242, 242, 242) : Color.Argb(255, 20, 20, 20);
        var faint = dark ? Color.Argb(255, 110, 110, 110) : Color.Argb(255, 140, 140, 140);
        var selected = dark ? Color.Argb(255, 36, 36, 36) : Color.Argb(255, 237, 237, 237);

        var bitmap = Bitmap.CreateBitmap(width, height, Bitmap.Config.Argb8888!)!;
        var canvas = new Canvas(bitmap);
        canvas.DrawColor(background);

        var settings = Workspace.Settings;
        var firstDay = settings.FirstDayOfWeek;
        var colorWeekends = settings.ColorWeekends;

        var month = new DateTime(today.Year, today.Month, 1);
        var lead = ((int)month.DayOfWeek - (int)firstDay + 7) % 7;
        var start = month.AddDays(-lead);
        var end = start.AddDays(Rows * Columns - 1);

        var withEvents = Workspace.Repository.DaysWithEvents(start, end);
        var withTodos = Workspace.Repository.OpenTodos()
            .Where(t => t.DueAt.HasValue)
            .Select(t => t.DueAt!.Value.Date)
            .ToHashSet();

        // ---- the measures, all from the shape given ----------------------------------
        // The title takes a band at the top, the headings a thinner one, and the six rows
        // share what is left. Text is sized from the cell it has to sit in, never from
        // the picture as a whole, so a wide widget does not get a title that dwarfs the
        // numbers and a tall one does not get numbers that dwarf the title.
        var pad = 0.02f * Math.Min(width, height);
        var titleBand = 0.13f * height;
        var headingBand = 0.07f * height;
        var gridTop = titleBand + headingBand;

        var cellWidth = width / (float)Columns;
        var cellHeight = (height - gridTop - pad) / Rows;
        var cell = Math.Min(cellWidth, cellHeight);

        var titleSize = Math.Min(titleBand * 0.55f, width * 0.07f);
        var headingSize = Math.Min(headingBand * 0.62f, cellWidth * 0.3f);
        var daySize = cell * 0.42f;
        var dotRadius = Math.Max(2f, daySize * 0.13f);
        var inset = cellWidth * 0.06f;
        var corner = cell * 0.14f;

        using var text = new Paint(PaintFlags.AntiAlias) { TextAlign = Paint.Align.Center };
        using var fill = new Paint(PaintFlags.AntiAlias) { };

        // ---- title -------------------------------------------------------------
        text.Color = primary;
        text.TextSize = titleSize;
        text.FakeBoldText = true;
        canvas.DrawText(today.ToString("yyyy년 M월"), width / 2f, Baseline(titleBand / 2f, text), text);

        // ---- weekday headings --------------------------------------------------
        text.TextSize = headingSize;
        text.FakeBoldText = false;
        var headingCentre = titleBand + (headingBand / 2f);
        for (var i = 0; i < Columns; i++)
        {
            var day = (DayOfWeek)(((int)firstDay + i) % 7);
            text.Color = WeekdayColor(day, faint, colorWeekends);
            canvas.DrawText(ShortName(day), (i + 0.5f) * cellWidth, Baseline(headingCentre, text), text);
        }

        // ---- the days ----------------------------------------------------------
        for (var index = 0; index < Rows * Columns; index++)
        {
            var date = start.AddDays(index);
            var column = index % Columns;
            var row = index / Columns;

            var centreX = (column + 0.5f) * cellWidth;
            var top = gridTop + (row * cellHeight);
            var centreY = top + (cellHeight / 2f);

            if (date.Date == today.Date)
            {
                fill.Color = selected;
                canvas.DrawRoundRect(
                    centreX - (cellWidth / 2f) + inset,
                    top + (inset / 2f),
                    centreX + (cellWidth / 2f) - inset,
                    top + cellHeight - (inset / 2f),
                    corner, corner, fill);
            }

            var inMonth = date.Month == month.Month && date.Year == month.Year;
            text.Color = inMonth
                ? WeekdayColor(date.DayOfWeek, primary, colorWeekends)
                : faint;
            text.TextSize = daySize;
            text.FakeBoldText = date.Date == today.Date;

            // The number sits a little above the middle of the cell to leave room for the
            // dots under it, so that number and dots together are centred.
            var numberCentre = centreY - (daySize * 0.22f);
            canvas.DrawText(date.Day.ToString(), centreX, Baseline(numberCentre, text), text);

            if (!inMonth) continue;

            // Two dots under the number, same meaning as in the app: solid for an event,
            // faint for a todo.
            var dotY = numberCentre + (daySize * 0.72f);
            var hasEvent = withEvents.Contains(date);
            var hasTodo = withTodos.Contains(date);
            var offset = hasEvent && hasTodo ? dotRadius * 1.8f : 0f;

            if (hasEvent)
            {
                fill.Color = primary;
                canvas.DrawCircle(centreX - offset, dotY, dotRadius, fill);
            }

            if (hasTodo)
            {
                fill.Color = faint;
                canvas.DrawCircle(centreX + offset, dotY, dotRadius, fill);
            }
        }

        return bitmap;
    }

    /// <summary>
    /// The baseline that puts the text's visual centre at <paramref name="centreY"/>.
    /// DrawText takes a baseline, and the distance from baseline to the middle of a digit
    /// depends on the size — which is no longer a constant here.
    /// </summary>
    private static float Baseline(float centreY, Paint text)
    {
        var metrics = text.GetFontMetrics()!;
        return centreY - ((metrics.Ascent + metrics.Descent) / 2f);
    }

    private static Color WeekdayColor(DayOfWeek day, Color fallback, bool colorWeekends)
    {
        if (!colorWeekends) return fallback;

        return day switch
        {
            DayOfWeek.Saturday => Color.Argb(255, 91, 141, 239),
            DayOfWeek.Sunday => Color.Argb(255, 224, 106, 106),
            _ => fallback,
        };
    }

    private static string ShortName(DayOfWeek day) => day switch
    {
        DayOfWeek.Monday => "월",
        DayOfWeek.Tuesday => "화",
        DayOfWeek.Wednesday => "수",
        DayOfWeek.Thursday => "목",
        DayOfWeek.Friday => "금",
        DayOfWeek.Saturday => "토",
        _ => "일",
    };
}
