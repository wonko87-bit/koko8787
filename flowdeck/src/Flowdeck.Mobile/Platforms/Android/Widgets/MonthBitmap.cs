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
/// </summary>
internal static class MonthBitmap
{
    private const int Size = 1000;
    private const int Columns = 7;
    private const int Rows = 6;

    public static Bitmap Render(DateTime today, bool dark)
    {
        var background = dark ? Color.Argb(235, 15, 15, 15) : Color.Argb(240, 255, 255, 255);
        var primary = dark ? Color.Argb(255, 242, 242, 242) : Color.Argb(255, 20, 20, 20);
        var faint = dark ? Color.Argb(255, 110, 110, 110) : Color.Argb(255, 140, 140, 140);
        var selected = dark ? Color.Argb(255, 36, 36, 36) : Color.Argb(255, 237, 237, 237);

        var bitmap = Bitmap.CreateBitmap(Size, Size, Bitmap.Config.Argb8888!)!;
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

        using var text = new Paint(PaintFlags.AntiAlias) { TextAlign = Paint.Align.Center };
        using var fill = new Paint(PaintFlags.AntiAlias) { };

        // ---- title -------------------------------------------------------------
        text.Color = primary;
        text.TextSize = 62;
        text.FakeBoldText = true;
        canvas.DrawText(today.ToString("yyyy년 M월"), Size / 2f, 74, text);

        // ---- weekday headings --------------------------------------------------
        const float gridTop = 130f;
        var cellWidth = Size / (float)Columns;
        var cellHeight = (Size - gridTop - 16) / Rows;

        text.TextSize = 34;
        text.FakeBoldText = false;
        for (var i = 0; i < Columns; i++)
        {
            var day = (DayOfWeek)(((int)firstDay + i) % 7);
            text.Color = WeekdayColor(day, faint, colorWeekends);
            canvas.DrawText(ShortName(day), (i + 0.5f) * cellWidth, gridTop - 20, text);
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
                    centreX - (cellWidth / 2f) + 8,
                    top + 4,
                    centreX + (cellWidth / 2f) - 8,
                    top + cellHeight - 4,
                    14, 14, fill);
            }

            var inMonth = date.Month == month.Month && date.Year == month.Year;
            text.Color = inMonth
                ? WeekdayColor(date.DayOfWeek, primary, colorWeekends)
                : faint;
            text.TextSize = 40;
            text.FakeBoldText = date.Date == today.Date;
            canvas.DrawText(date.Day.ToString(), centreX, centreY + 2, text);

            if (!inMonth) continue;

            // Two dots under the number, same meaning as in the app: solid for an event,
            // faint for a todo.
            var dotY = centreY + 28;
            var hasEvent = withEvents.Contains(date);
            var hasTodo = withTodos.Contains(date);
            var offset = hasEvent && hasTodo ? 9f : 0f;

            if (hasEvent)
            {
                fill.Color = primary;
                canvas.DrawCircle(centreX - offset, dotY, 5, fill);
            }

            if (hasTodo)
            {
                fill.Color = faint;
                canvas.DrawCircle(centreX + offset, dotY, 5, fill);
            }
        }

        return bitmap;
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
