using Microsoft.Graph;

static class TimeCardDuration
{
    public static string FormatDuration(TimeSpan? ts)
    {
        if (ts == null)
            return "--:--:--";
        var tss = (TimeSpan)ts;
        return $"{(tss.Days * 24) + tss.Hours:00}:{tss.Minutes:00}:{tss.Seconds:00}";
    }
    public static string FormatDurationShort(TimeSpan? ts)
    {
        if (ts == null)
            return "--:--";
        var tss = (TimeSpan)ts;
        return $"{(tss.Days * 24) + tss.Hours:00}:{tss.Minutes:00}";
    }

    public static TimeSpan? TotalTimespan(IEnumerable<TimeCard>? timeCards)
    {
        if (timeCards == null)
            return null;

        var total = new TimeSpan();
        foreach (var tc in timeCards)
        {
            total += HoursFromTimeCard(tc);
        }

        // This is required to stop blazor from infinitely re-rendering the page
        return new TimeSpan(total.Days, total.Hours, total.Minutes, total.Seconds);
    }


    public static TimeSpan HoursFromTimeCard(TimeCard tc)
    {
            DateTime dt_in = tc!.ClockInEvent?.DateTime?.DateTime ?? throw new Exception("TimeCard.ClockInEvent.DateTime should never not be null");
            DateTime dt_out = tc!.ClockOutEvent?.DateTime?.DateTime ?? DateTime.UtcNow; // Still clocked in
            return dt_out - dt_in - BreakHoursFromTimeCard(tc);
    }
    public static TimeSpan BreakHoursFromTimeCard(TimeCard tc)
    {
        var total = new TimeSpan();
        foreach (var b in tc?.Breaks ?? Enumerable.Empty<TimeCardBreak>())
            total += HoursFromTimeCardBreak(b);
        return total;
    }

    public static  TimeSpan HoursFromTimeCardBreak(TimeCardBreak tcb)
    {
        DateTime dt_in = tcb!.Start?.DateTime?.DateTime ?? throw new Exception("TimeCardBreak.Start.DateTime should never not be null");
        DateTime dt_out = tcb?.End?.DateTime?.DateTime ?? DateTime.UtcNow; // Still on break
        var hours = (dt_out - dt_in);
        return hours;
    }

}
