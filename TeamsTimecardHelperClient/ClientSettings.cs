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

    public DateOnly statsstartdate {
        get {
            string date_stamp = _localstorage.GetItem<string>("statsstartdate");
            try {
                return DateOnly.FromDateTime(DateTime.Parse(date_stamp));
            }
            catch (Exception)
            {
                return new DateOnly(DateTime.Now.Year, 1, 1);
            }
        }
        set {
            string date_stamp = value.ToString();
            _localstorage.SetItem("statsstartdate", date_stamp);
        }
    }

    public int targetweeklyhours {
        get => _localstorage.GetItem<int>("targetweeklyhours");
        set => _localstorage.SetItem("targetweeklyhours", value);
    }
}
