using Blazored.LocalStorage;
internal class ClientSettings
{
    private static ISyncLocalStorageService _localstorage = null!;
    public ClientSettings(ISyncLocalStorageService localstorage)
    {
        _localstorage = localstorage;
    }

    public string? theteamid {
        get => _localstorage.GetItem<string>("theteamid");
        set =>_localstorage.SetItem("theteamid", value);
    }
}
