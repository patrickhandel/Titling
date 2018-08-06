using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Threading.Tasks;
using Geocoding;
using Geocoding.Google;
using System.Linq;

namespace DOT_Titling_Excel_VSTO
{
    class Geocode
    {
        public async static Task<bool> ExecuteGeocode(Excel.Application app)
        {
            try
            {
                Excel.Worksheet ws = app.ActiveSheet;
                Excel.Range activeCell = app.ActiveCell;
                Excel.Range selection = app.Selection;


                //for (int row = selection.Row; row < selection.Row + selection.Rows.Count; row++)
                //{
                //    if (ws.Rows[row].EntireRow.Height != 0)
                //    {

                //    }
                //}

                //AIzaSyBU7taGZyxfMhiO99c8F40OznPiDFN9LR4

                IGeocoder geocoder = new GoogleGeocoder() { ApiKey = "AIzaSyBU7taGZyxfMhiO99c8F40OznPiDFN9LR4" };
                IEnumerable<Address> addresses = await geocoder.GeocodeAsync("1600 pennsylvania ave washington dc");
                Console.WriteLine("Formatted: " + addresses.First().FormattedAddress); //Formatted: 1600 Pennsylvania Ave SE, Washington, DC 20003, USA
                Console.WriteLine("Coordinates: " + addresses.First().Coordinates.Latitude + ", " + addresses.First().Coordinates.Longitude); //Coordinates: 38.8791981, -76.9818437
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error :" + ex);
                return false;
            }
        }
    }
}
