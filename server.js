var express = require("express");
var app = express();
var xlsxtojson = require("xlsx-to-json");
var bodyParser = require("body-parser");
var json2xls = require("json2xls");

const fs = require("fs");

// r=r.concat(j);
app.use(function (req, res, next) {
  //allow cross origin requests
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "POST, PUT, OPTIONS, DELETE, GET");
  res.header("Access-Control-Max-Age", "3600");
  res.header(
    "Access-Control-Allow-Headers",
    "Content-Type, Access-Control-Allow-Headers, Authorization, X-Requested-With"
  );
  next();
});

// configuration
app.use(express.static(__dirname + "/public"));
app.use("/public/uploads", express.static(__dirname + "/public/uploads"));
app.use(bodyParser());
app.use(json2xls.middleware);
app.get("/", function (req, res) {
  res.send("Hello World");
});
//POST REQUEST

// var urlencodedParser = bodyParser.urlencoded({ extended: false });
app.get("/excel", function (req, res) {
  res.send("Hello World");
});
app.post("/excel", function (req, res,next) {
  var r = require("./output.json");
  res.end(JSON.stringify(req.body));
  var incoming = req.body;
  r = r.concat(incoming);
  console.log(r);
  var k = r;
  var xls = json2xls(k);
  fs.writeFileSync("./excel-to-json.xlsx", xls, "binary");
  //   res.end(JSON.stringify(req.body));
  console.log(req.body);
  next();
});

//GET REQUEST
app.get("/api/data", function (req, res) {
  res.redirect("/api/xlstojson");
});
app.get("/api/xlstojson", function (req, res) {
  xlsxtojson(
    {
      input: "./excel-to-json.xlsx", // input xls
      output: "output.json", // output json
      lowerCaseHeaders: true,
    },
    function (err, result) {
      if (err) {
        res.json(err);
      } else {
        res.json(result);
      }
    }
  );
});

app.get("/api/data", function (req, res) {
  res.send(app.locals.data);
});

app.listen(5000);
