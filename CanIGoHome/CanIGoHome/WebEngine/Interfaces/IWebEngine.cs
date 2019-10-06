using System.Collections.Generic;

namespace CanIGoHome
{
    public interface IWebEngine
    {
        bool ConfigureEngine(IDictionary<string,string> settings);

        string Search();
    }
}
