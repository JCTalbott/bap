using System.Globalization;

public class ExcelDataWorker {
    public ExcelDataWorker() {

    }

    public string dateWorker(string daysFrom1900) {
        string dateFormat = "";
        int daysInt = 0;
        DateTime myDT = new DateTime( 1900, 1, 1, new GregorianCalendar() );
        Calendar myCal = CultureInfo.InvariantCulture.Calendar;
        if (Int32.TryParse(daysFrom1900, out daysInt))
        {
            myDT = myCal.AddDays( myDT, daysInt );
            string year = myCal.GetYear( myDT ).ToString();
            string month = myCal.GetMonth( myDT ).ToString();
            string dayOfMonth = (myCal.GetDayOfMonth( myDT ) - 1).ToString(); // minus 1 idk why
            dateFormat = $"{month}/{dayOfMonth}/{year}";
        }
        return dateFormat;
    }

    public string currencyWorker(string money) {
        double m;
        if (double.TryParse(money, out double d)) {
            return d.ToString("C", CultureInfo.CurrentCulture);
        }
        return money;
    }

    public string percentageWorker(string percent) {
        double p;
        if (double.TryParse(percent, out p)) {
            return p.ToString("P1", CultureInfo.InvariantCulture);
        }
        return percent;
    }
}