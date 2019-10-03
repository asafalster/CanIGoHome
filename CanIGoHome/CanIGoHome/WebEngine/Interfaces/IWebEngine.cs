namespace CanIGoHome
{
    public interface IWebEngine
    {
        bool ConfigureEngine();

        string Search(string user, string pw);
    }
}
