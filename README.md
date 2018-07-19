This project was forked from https://github.com/michaelneu/webxcel.git and was adapted for my tasks

I removed some functionality from the parent project e.g. primary and foreign keys to simplify work with tables.
And I changed an algorithm working of HttpServer because parent project's HttpServer doesn't allow use Formulas. 

Webxcel creates a full-fledged RESTful web backend from your Microsoft Excel workbooks and you can use excel functionality to calculation any values.

To create new service, simply create new sheet, insert your column names in the first row of the sheet and write write formulas in the second row. E.g. : 

When accessing Post /fio
HTTP/1.1 200 Nobody Needs This Anyway
Content-Type: application/json
Server: Microsoft Excel/16.0
Content-Length: 200
[
  {
    "id": "1",
    "FirstName": "Andrey",
    "LastName":"T",
  },
  {
    "id": "1",
    "FirstName": "Mark",
    "LastName":"F",
  }
]
Response:
[
  {
    "id": "1",
    "FirstName": "Andrey",
    "LastName":"T",
    "FIO" : "Andrey T"
  },
  {
    "id": "1",
    "FirstName": "Mark",
    "LastName":"F",
    "FIO" : "Mark F"
  }
]

Using mechanism of Application.OnTime allows to work with addons formulas which can't finish calculation if you use simple loop ( e.g. Bloomberg) but sometimes requests aren't handled because calling Application.Ontime is interruption and request fall into time interval of interruption but it happens with 	low-probability



## License

Webxcel is released under the [MIT license](LICENSE).
