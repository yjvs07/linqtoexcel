# Linq to Excel #

Use LINQ to retrieve data from spreadsheets.
```
var excel = new ExcelQueryFactory("excelFileName");
var indianaCompanies = from c in excel.Worksheet<Company>()
                       where c.State == "IN"
                       select c;
```

---

The home page is now located on [Github](https://github.com/paulyoder/LinqToExcel)

Install the [NuGet package](http://nuget.org/List/Packages/LinqToExcel).

Go to the [Read me](https://github.com/paulyoder/LinqToExcel#welcome-to-the-linqtoexcel-project) page for information on implementing Linq to Excel in your project.

Need help? Report an [issue](https://github.com/paulyoder/LinqToExcel/issues/new) or ask a question on Stack Overflow.

---

### Demo Video ###
<a href='http://www.youtube.com/watch?feature=player_embedded&v=t3BEUP0OTFM' target='_blank'><img src='http://img.youtube.com/vi/t3BEUP0OTFM/0.jpg' width='640' height=385 /></a>