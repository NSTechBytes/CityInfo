/*
       
//=====================================================================================================================================//
//                              Below Lines Help to use Plugin with Measures                                                           //
//=====================================================================================================================================//

[Searcher_Helper]
Measure=Plugin
Plugin=CityInfo
ResultsSave=#@#Includes\SearchResults.nek
ApiKey=#OpenCage_ApiKey#
OnCompleteAction=[!Delay 100][!CommandMeasure Func "startResults('BackGround_Shape','1','110')"]


For Execution ([!CommandMeasure Searcher_Helper " City name"])
  */
//=====================================================================================================================================//
//                                             Main  Code Start Here                                                                   //
//=====================================================================================================================================//

using System;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Rainmeter;

namespace CityInfo
{
    internal class Measure
    {
        private string ApiKey;
        private string ResultSave;
        private string OnCompleteAction;
        private API api;
        private string LastCity = string.Empty;

        internal Measure() { }

        internal void Reload(API api, ref double maxValue)
        {
            this.api = api;
            ApiKey = api.ReadString("ApiKey", "").Trim();
            ResultSave = api.ReadString("ResultsSave", "").Trim();
            OnCompleteAction = api.ReadString("OnCompleteAction", "").Trim();

            if (string.IsNullOrEmpty(ApiKey))
            {
                api.Log(API.LogType.Error, "CityInfo.dll: 'ApiKey' must be provided.");
            }
            if (string.IsNullOrEmpty(ResultSave))
            {
                api.Log(API.LogType.Error, "CityInfo.dll: 'ResultSave' must be provided.");
            }
        }

        internal void Execute(string args)
        {
            if (string.IsNullOrEmpty(args))
            {
                api.Log(API.LogType.Error, "CityInfo.dll: No city name provided.");
                return;
            }


            LastCity = args.Replace("Execute", "").Trim();

            string url =
                $"https://api.opencagedata.com/geocode/v1/json?q={Uri.EscapeDataString(LastCity)}&key={ApiKey}";
            string response = FetchCityData(url);

            if (response == null)
            {
                SaveErrorResult("Fail to fetch data. Make sure you are connected to the internet.");
            }
            else if (!ProcessApiResponse(response))
            {

                SaveErrorResult($"No result found for \"{LastCity}\".");
            }

            if (!string.IsNullOrEmpty(OnCompleteAction))
            {
                api.Execute(OnCompleteAction);
            }
        }

        private string FetchCityData(string url)
        {
            try
            {

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "GET";

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    api.Log(
                        API.LogType.Debug,
                        $"CityInfo.dll: Response status code: {response.StatusCode}"
                    );
                    using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            catch (WebException ex)
            {
                api.Log(API.LogType.Error, $"CityInfo.dll: HTTP request failed: {ex.Message}");
                return null;
            }
        }

        private bool ProcessApiResponse(string jsonResponse)
        {
            string inputCityName = LastCity;
            string fetchedCityName = string.Empty;
            string country = string.Empty;
            string longitude = string.Empty;
            string latitude = string.Empty;

            try
            {
                if (jsonResponse.Contains("\"results\":[]"))
                {
                    return false;
                }


                string geometryPattern = @"\""geometry\"":\s*{[^}]*}";
                string latPattern = @"\""lat\"":\s*(-?\d+(\.\d+)?)";
                string lngPattern = @"\""lng\"":\s*(-?\d+(\.\d+)?)";
                string countryPattern = @"\""country\"":\s*\""([^\""]+)\""";
                string cityPattern = @"\""city\"":\s*\""([^\""]+)\""";


                Match geometryMatch = Regex.Match(jsonResponse, geometryPattern);
                if (geometryMatch.Success)
                {
                    string geometryBlock = geometryMatch.Value;


                    Match latMatch = Regex.Match(geometryBlock, latPattern);
                    if (latMatch.Success)
                    {
                        latitude = latMatch.Groups[1].Value;
                    }


                    Match lngMatch = Regex.Match(geometryBlock, lngPattern);
                    if (lngMatch.Success)
                    {
                        longitude = lngMatch.Groups[1].Value;
                    }
                }


                Match countryMatch = Regex.Match(jsonResponse, countryPattern);
                if (countryMatch.Success)
                {
                    country = countryMatch.Groups[1].Value;
                }


                Match cityMatch = Regex.Match(jsonResponse, cityPattern);
                if (cityMatch.Success)
                {
                    fetchedCityName = cityMatch.Groups[1].Value;
                }


                if (string.IsNullOrEmpty(fetchedCityName))
                {
                    fetchedCityName = inputCityName;
                }


                fetchedCityName = fetchedCityName.Replace("Execute", "").Trim();


                using (StreamWriter writer = new StreamWriter(ResultSave, false))
                {
                    writer.WriteLine("[Result_1]");
                    writer.WriteLine("Meter=String");
                    writer.WriteLine($"Text={fetchedCityName}, {country}");
                    writer.WriteLine("MeterStyle=Result_String");
                    writer.WriteLine(
                        $"LeftMouseUpAction=[!WriteKeyValue Variables Longitude \"{longitude}\" \"#@#GlobalVar.nek\"]"
                            + $"[!WriteKeyValue Variables Latitude \"{latitude}\" \"#@#GlobalVar.nek\"]"
                            + $"[!WriteKeyValue Variables City \"{fetchedCityName}\" \"#@#GlobalVar.nek\"][!WriteKeyValue Variables Country \"{country}\" \"#@#GlobalVar.nek\"][!SetVariable Longitude  \"{longitude}\" \"#NekStart\\Main\"][!SetVariable Latitude  \"{latitude}\" \"#NekStart\\Main\"][!SetVariable City  \"{fetchedCityName}\" \"#NekStart\\Main\"][!SetVariable Country  \"{country}\" \"#NekStart\\Main\"][!UpdateMeter \"*\" \"#NekStart\\Main\"][!Redraw \"#NekStart\\Main\"][!UpdateMeasure mToggle]"
                    );
                }

                api.Log(
                    API.LogType.Debug,
                    $"CityInfo.dll: Data successfully saved for city: {fetchedCityName}"
                );
                return true;
            }
            catch (Exception ex)
            {
                api.Log(
                    API.LogType.Error,
                    $"CityInfo.dll: Error processing API response: {ex.Message}"
                );
                return false;
            }
        }



        private void SaveErrorResult(string errorMessage)
        {
            using (StreamWriter writer = new StreamWriter(ResultSave, false))
            {
                writer.WriteLine("[Result_1]");
                writer.WriteLine("Meter=String");
                writer.WriteLine($"Text={errorMessage}");
                writer.WriteLine("MeterStyle=Result_String");
                writer.WriteLine("MouseOverAction=[]");
                writer.WriteLine("MouseLeaveAction=[]");
            }
        }
    }
    //=====================================================================================================================================//
    //                                            Rainmeter Class                                                                          //
    //=====================================================================================================================================//
    public static class Plugin
    {
        [DllExport]
        public static void Initialize(ref IntPtr data, IntPtr rm)
        {
            data = GCHandle.ToIntPtr(GCHandle.Alloc(new Measure()));
        }

        [DllExport]
        public static void Finalize(IntPtr data)
        {
            GCHandle.FromIntPtr(data).Free();
        }

        [DllExport]
        public static void Reload(IntPtr data, IntPtr rm, ref double maxValue)
        {
            Measure measure = (Measure)GCHandle.FromIntPtr(data).Target;
            measure.Reload(new API(rm), ref maxValue);
        }

        [DllExport]
        public static double Update(IntPtr data)
        {
            return 0.0;
        }

        [DllExport]
        public static void ExecuteBang(IntPtr data, [MarshalAs(UnmanagedType.LPWStr)] string args)
        {
            Measure measure = (Measure)GCHandle.FromIntPtr(data).Target;
            measure.Execute(args);
        }
    }
}
