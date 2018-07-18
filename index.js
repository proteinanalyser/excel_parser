var express = require('express');
var bodyParser = require("body-parser");
var path = require("path");
var app = express();
var multer  = require('multer')
var XLSX = require('xlsx')

const storage = multer.diskStorage({
  destination: "./uploads",
  filename: function (req, file, cb) {
    cb(null, file.originalname);
  }
});
const upload = multer({
  storage: storage,
  fileFilter: function (req, file, cb) {
    checkFileType(file, cb);
  }
}).single("myfile");

function checkFileType(file, cb) {
  let filetypes = ".xlsx";
  let extname = path.extname(file.originalname).toLowerCase() == filetypes;
  let mimetype = file.mimetype == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  if (originalname != ".xlsx") {
    cb("The file you uploaded has a wrong file extension! The file must be .xlsx");
  } else if(file.name !="template.xlsx"){
    cb("Only upload the template file provided by this platform!");
  } else if ( XLSX.read("./uploads/template.xlsx", {type:'file'}).SheetNames.length > 1){
    cb("Please do not add any excel sheets to the template.");
  }else if ( XLSX.read("./uploads/template.xlsx", {type:'file'}).SheetNames[0] != "Sheet1"){
    cb("Please do not change the name of the excel sheet");
  } else if (mimetype && extname) {
    return cb(null, true);
  }
}
// Add headers
app.use(function (req, res, next) {

  // Website you wish to allow to connect
  res.setHeader('Access-Control-Allow-Origin', 'http://agirar.pt');
  // res.setHeader('Access-Control-Allow-Origin', 'http://localhost:8000');

  // Request methods you wish to allow
  res.setHeader('Access-Control-Allow-Methods', 'POST');

  // Request headers you wish to allow
  res.setHeader('Access-Control-Allow-Headers', 'X-Requested-With,content-type');
  res.setHeader('Access-Control-Allow-Credentials', true);

  // Pass to next layer of middleware
  next();
});

app.set("port", (process.env.PORT || 3000));

// Body Parser Middleware
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({
  extended: false
}));

// Get Static Path
app.use(express.static(path.join(__dirname, "public")));


app.post("/parse_excel", function(req, res){
  upload(req, res, (err) => {
    if (err) {
      res.send({
        error: true,
        msg: err
      });
    } else {
      var workbook = XLSX.read("./uploads/template.xlsx", {type:'file'});
      var first_sheet_name = workbook.SheetNames[0];
      var worksheet = workbook.Sheets[first_sheet_name];
      
      res.send({
        error: false,
        data: XLSX.utils.sheet_to_json(worksheet)
      })
    }
  });
})
// 
app.listen(app.get("port"), function() {
  console.log("App is listenning o port 3000");
});
