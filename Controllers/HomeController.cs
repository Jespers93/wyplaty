using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.IO;
using wyplaty.Entities;
using System.Globalization;
using Microsoft.EntityFrameworkCore;

namespace wyplaty.Controllers
{
    [Route("api/[controller]")]
    public class HomeController : ControllerBase
    {
        private readonly FrachtDbContext _dbcontext;

        public HomeController(FrachtDbContext dbcontext)
        {
            _dbcontext = dbcontext;
        }


        [HttpGet("again")]
        public async Task<ActionResult<int>> GetFrachtAgain()
        {

            FileStream fs = new FileStream("C:\\Users\\Krzysztof.KWiatkowsk\\source\\repos\\wyplaty\\Files\\2023listopad-maciej.xlsx", FileMode.Open,
    FileAccess.ReadWrite, FileShare.None, 1024);

            string test = "";
            int count;
            using (StreamReader sr = new StreamReader(fs))
            {
                ExcelPackage pkg = new ExcelPackage(fs);
                count = pkg.Workbook.Worksheets.Count;

                for (int i = 0; i < count; i++)
                {

                

                var ws = pkg.Workbook.Worksheets[i];
                var totalcount = ws.Dimension.End.Row;
                for (int y = 2; y <= totalcount; y++)
                {

                        Fracht frachtemp = new Fracht();
                        frachtemp.OrderNumber = ws.Cells[y,1] + ws.Cells[y, 5].Text.Trim() + ws.Cells[y, 6].Text.Replace(" ", "") + ws.Cells[y, 7].Text;
                        test = frachtemp.OrderNumber;
                        if (!_dbcontext.Frachts.Any(f => f.OrderNumber == frachtemp.OrderNumber) && frachtemp.OrderNumber != "")
                        {
                            var datetest = ws.Cells[y, 3].Value;
                            DateTime loadDate = new DateTime();
                            if (datetest.GetType() == typeof(DateTime))
                            {
                                loadDate = (DateTime)datetest;
                            }
                            else { 
                                long dateLoadNum = long.Parse(ws.Cells[y, 3].Value.ToString());
                                loadDate = DateTime.FromOADate(dateLoadNum);
                            }
                            frachtemp.LoadDate = loadDate;



                            var datetestun = ws.Cells[y, 5].Value;
                            DateTime unloadDate = new DateTime();
                            if (datetestun.GetType() == typeof(DateTime))
                            {
                                unloadDate = (DateTime)datetestun;
                            }
                            else
                            {
                                long dateUnloadNum = long.Parse(ws.Cells[y, 5].Value.ToString());
                                unloadDate = DateTime.FromOADate(dateUnloadNum);
                            }                            
                            frachtemp.UnloadDate = unloadDate;


                            frachtemp.Car = ws.Cells[y,2].Text.Replace(" ","");
                            frachtemp.CityLoad = ws.Cells[y, 4].Text;
                            frachtemp.CityUnload = ws.Cells[y, 6].Text;

                            if (ws.Cells[y, 7].Text != "")
                                frachtemp.Price = double.Parse(ws.Cells[y, 7].Text);
                            else frachtemp.Price = 0;

                            var lastnametemp = ws.Cells[y, 1].Text;
                            var drivercount = _dbcontext.Drivers.Where(x => x.LastName == lastnametemp).Count();

                            if (drivercount == 0)
                            {

                                Driver drivertemp = new Driver
                                {
                                    LastName = lastnametemp
                                };
                                _dbcontext.Drivers.Add(drivertemp);
                                frachtemp.Driver = drivertemp;
                            }
                            else
                                frachtemp.Driver = _dbcontext.Drivers.FirstOrDefault(x => x.LastName == lastnametemp);

                            await _dbcontext.Frachts.AddAsync(frachtemp);
                            await _dbcontext.SaveChangesAsync();

                        }
                }

                }
            }

                return Ok(test);
        }

        [HttpGet]
        public async Task<ActionResult> GetFracht()
        {


            FileStream fs = new FileStream("C:\\Users\\Krzysztof.KWiatkowsk\\source\\repos\\wyplaty\\Files\\2023listopad-martrans.xlsx", FileMode.Open,
    FileAccess.ReadWrite, FileShare.None, 1024);


            //FileInfo newFile = new FileInfo("C:\\Users\\Krzysztof.KWiatkowsk\\source\\repos\\wyplaty\\Files\\test.xlsx");

            using (StreamReader sr = new StreamReader(fs))
            {
                ExcelPackage pkg = new ExcelPackage(fs);
                var ws = pkg.Workbook.Worksheets.First();

                var totalcount = ws.Dimension.End.Row;

                var test = ws.Cells[2, 1].Value;

                

                for (int i = 2; i<= totalcount; i++)
                {
                    Fracht frachtemp = new Fracht();
                    frachtemp.OrderNumber = ws.Cells[i,1].Text;
                    if (!_dbcontext.Frachts.Any(o => o.OrderNumber == frachtemp.OrderNumber))
                    {
                        long dateLoadNum = long.Parse(ws.Cells[i,2].Value.ToString());
                        DateTime loadDate = DateTime.FromOADate(dateLoadNum);
                        frachtemp.LoadDate = loadDate;

                        long dateUnloadNum = long.Parse(ws.Cells[i,3].Value.ToString());
                        DateTime unloadDate = DateTime.FromOADate(dateUnloadNum);
                        frachtemp.UnloadDate = unloadDate;

                        frachtemp.Car = ws.Cells[i,4].Value.ToString();
                        frachtemp.CityLoad = ws.Cells[i,6].Value.ToString();
                        frachtemp.CityUnload = ws.Cells[i,7].Value.ToString();
                        frachtemp.Price = float.Parse(ws.Cells[i,9].Text);

                        if (ws.Cells[i, 10].Text != "")
                        {

                            var lastnametemp = ws.Cells[i, 10].Value.ToString();
                            var drivercount = _dbcontext.Drivers.Where(x => x.LastName == lastnametemp).Count();

                            if (drivercount == 0)
                            {

                                Driver drivertemp = new Driver
                                {
                                    LastName = lastnametemp
                                };
                                _dbcontext.Drivers.Add(drivertemp);
                                frachtemp.Driver = drivertemp;
                            }
                            else
                                frachtemp.Driver = _dbcontext.Drivers.FirstOrDefault(x => x.LastName == lastnametemp);
                        }
                        else throw new Exception("Wystąpił błąd, bo nie ma kierowccy");
                        await _dbcontext.Frachts.AddAsync(frachtemp);
                        await _dbcontext.SaveChangesAsync();
                    }
                }

                /*  frachtemp.OrderNumber = ws.Cells["A2"].Text;

                  long dateLoadNum = long.Parse(ws.Cells["B2"].Value.ToString());
                  DateTime loadDate = DateTime.FromOADate(dateLoadNum);
                  frachtemp.LoadDate = loadDate;

                  long dateUnloadNum = long.Parse(ws.Cells["C2"].Value.ToString());
                  DateTime unloadDate = DateTime.FromOADate(dateUnloadNum);
                  frachtemp.UnloadDate = unloadDate;

                  frachtemp.Car = ws.Cells["D2"].Value.ToString();
                  frachtemp.CityLoad = ws.Cells["G2"].Value.ToString();
                  frachtemp.CityUnload = ws.Cells["J2"].Value.ToString();
                  frachtemp.Price = int.Parse(ws.Cells["L2"].Text);
  */

                pkg.Save();
            }

            /*  if (_dbcontext.Frachts.Any(o => o.OrderNumber == frachtemp.OrderNumber))
              {
                  return NotFound();
              }
              else
              {
                  _dbcontext.Frachts.Add(frachtemp);
                  _dbcontext.SaveChanges();

                  return Ok(frachtemp);
              }
  */
            return Ok();
        }

        [HttpGet ("{car}")]
        public async Task<ActionResult<ICollection<Fracht>>> GetFrachtByCar (string car)
        {
            // return Ok(_dbcontext.Frachts.Where( o => o.Car == car).Sum(i => i.Price));
            return Ok(_dbcontext.Frachts.Where(o => o.Car == car));
        }

        [HttpGet ("copy/{car}")]
        public async Task<ActionResult<ICollection<string>>> GetFrachtByCarCopy(string car)
        {
            var test = _dbcontext.Frachts.Where(o => o.Car == car);

            List<string> test2 = new List<string>();

            foreach (var item in test)
            {

                test2.Add($"{item.Car} {item.CityLoad} - {item.CityUnload} \t {item.Price}");

            }
            return Ok(test2);
        }


        [HttpGet("copy/{car}/{month}/{year}")]
        public async Task<ActionResult<ICollection<string>>> GetFrachtByCarCopy(string car, int month, int year)
        {
            var test = _dbcontext.Frachts.Where(o => o.Car == car && o.UnloadDate.Month == month && o.UnloadDate.Year == year);

            List<string> test2 = new List<string>();

            foreach (var item in test)
            {

                test2.Add($"{item.Car} {item.CityLoad} - {item.CityUnload} \t {item.Price}");

            }
            return Ok(test2);
        }

        [HttpGet("copyall/{month}/{year}")]
        public async Task<ActionResult<ICollection<string>>> GetFrachtCopy(int month, int year)
        {
            var test = _dbcontext.Frachts.Where(o => o.UnloadDate.Month == month && o.UnloadDate.Year == year).OrderBy(z => z.Car).ThenBy(z=> z.UnloadDate);

            List<string> test2 = new List<string>();

            foreach (var item in test)
            {
               test2.Add($"{item.Car} {item.CityLoad} - {item.CityUnload} \t {item.Price} \t {item.UnloadDate.DayOfWeek} \t {ISOWeek.GetWeekOfYear(item.UnloadDate)}");

            }
            return Ok(test2);
        }

        [HttpGet("exportall/{month}/{year}")]
        public async Task<ActionResult> exportexcel(int month, int year)
        {

            

            FileInfo fs = new FileInfo("C:\\Users\\Krzysztof.KWiatkowsk\\source\\repos\\wyplaty\\Files\\wzor.xlsx");

            fs.CopyTo($"C:\\Users\\Krzysztof.KWiatkowsk\\source\\repos\\wyplaty\\Files\\wyplaty\\{month}-{year}.xlsx");

            fs = new FileInfo($"C:\\Users\\Krzysztof.KWiatkowsk\\source\\repos\\wyplaty\\Files\\wyplaty\\{month}-{year}.xlsx");

            ExcelPackage pkg = new ExcelPackage(fs);
                var wzorws = pkg.Workbook.Worksheets.First();

            var frachts = _dbcontext
                .Frachts
                .Include(x=> x.Driver)
                .Where(o => o.UnloadDate.Month == month && o.UnloadDate.Year == year)
                .OrderBy(o=> o.Car).ThenBy(z=>z.UnloadDate);

                foreach(var fracht in frachts)
            {
                var ws = pkg.Workbook.Worksheets[fracht.Driver.LastName];
                if (ws == null)
                {
                    ws = pkg.Workbook.Worksheets.Add(fracht.Driver.LastName, wzorws);
                    ws.Cells[2, 1].Value = "Wynagrodzenie " + month + "."+year;
                    ws.Cells[4, 2].Value = "01.11 - 05.11";
                    ws.Cells[16, 2].Value = "06.11 - 12.11";
                    ws.Cells[28, 2].Value = "13.11 - 19.11";
                    ws.Cells[40, 2].Value = "20.11 - 26.11";
                    ws.Cells[52, 2].Value = "27.11 - 30.11";
                }

                //  DateTime date = new DateTime(2022, 12, 16);
                //   int weekOfMonth2 = (date.Day + ((int)date.DayOfWeek)) / 7 + 1;

                var week1 = ISOWeek.GetWeekOfYear(fracht.UnloadDate);
                var week2 = ISOWeek.GetWeekOfYear(new DateTime(fracht.UnloadDate.Year, fracht.UnloadDate.Month, 1));

                if (week2 > week1) week2 = 1;

                int weekOfMonth = week1 - week2 + 1;

               if (weekOfMonth >= 6) weekOfMonth = 5;


                var indexwc = 4 + ((weekOfMonth - 1) * 12);


                while (ws.Cells[indexwc, 3].Value != null)
                {
                    indexwc++;
                    
                }


               



                if (ws.Cells[indexwc, 3].Value == null)
                {
                    ws.Cells[indexwc, 3].Value = fracht.UnloadDate.ToShortDateString();
                    ws.Cells[indexwc, 4].Value = ($"{fracht.Car} {fracht.CityLoad} - {fracht.CityUnload}");
                    ws.Cells[indexwc, 9].Value = fracht.Price;
                }
        
            }


                pkg.Save();

                return Ok();
        }


        [HttpPost]
        public async Task<ActionResult<string>> PostFracht([FromBody]string temp)
        {
            return Ok(temp);
        }
    }
}
