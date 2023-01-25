using CurrieTechnologies.Razor.PageVisibility;

internal static class RunIfVisible
{
    public static async Task RunIfVisibleAsync(this PageVisibilityService visibility, Func<Task> action)
    {
        var vis = await visibility.GetVisibilityStateAsync();
        if (vis == VisibilityState.Hidden)
            return;
        else
            await action();
    }
}