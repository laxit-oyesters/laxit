const express = require('express');
const mongoose = require('mongoose');
const bodyParser = require('body-parser');
const path = require('path');
const XLSX = require('xlsx');
const multer = require('multer');
const swaggerJsDoc = require("swagger-jsdoc");
const swaggerUi = require("swagger-ui-express");


var port = process.env.PORT || 3000;
//init app
var app = express();

const swaggerOptions = {
    swaggerDefinition: {
        info: {
            version: "1.0.0",
            title: "Customer API",
            description: "Customer API Information",
            contact: {
                name: "Amazing Developer"
            },
            servers: ["https://login-signup-register.herokuapp.com/"]
        }
    },
    // ['.routes/*.js']
    apis: ["server.js"]
};

const swaggerDocs = swaggerJsDoc(swaggerOptions);
app.use("/api-docs", swaggerUi.serve, swaggerUi.setup(swaggerDocs));

var storage = multer.diskStorage({
    destination: function (req, file, cb) {
      cb(null, './public/uploads')
    },
    filename: function (req, file, cb) {
      cb(null, file.originalname)
    }
  });
  
var upload = multer({ storage: storage });

//connect to db
mongoose.connect('mongodb://localhost:27017/Demoexcel',{useNewUrlParser:true})
.then(()=>{console.log('connected to db')})
.catch((error)=>{console.log('error',error)});


//set the template engine
app.set('view engine','ejs');

//fetch data from the request
app.use(bodyParser.urlencoded({extended:false}));

//static folder path
app.use(express.static(path.resolve(__dirname,'public')));

//collection schema
var excelSchema = new mongoose.Schema({
    name:String,
    Function:String,
    SubFunction:String,
    DepartmentLead:String,
    FunctionLead:String,
    SubFunctionLead:String,
    subject:String
});

var excelModel = mongoose.model('excelData',excelSchema);


app.get('/',(req,res)=>{
   excelModel.find((err,data)=>{
       if(err){
           console.log(err)
       }else{
           if(data!=''){
               res.render('home',{result:data});
           }else{
               res.render('home',{result:{}});
           }
       }
   });
});

app.post('/upload',upload.single('excel'),(req,res)=>{
  var workbook =  XLSX.readFile(req.file.path);
  var sheet_namelist = workbook.SheetNames;


  var sheet1 = workbook.Sheets["Sheet1"]
  var datas = XLSX.utils.sheet_to_json(sheet1)

  var newdata = datas.map(function(record){
    delete record.Function;
    return datas
  })

  console.log(newdata)



  newdata = newdata[0];
  console.log(newdata)
  var workbook = XLSX.utils.book_new();
  var ws = XLSX.utils.json_to_sheet(newdata)
  

  
  XLSX.utils.book_append_sheet(workbook,ws,"New Data")
  XLSX.writeFile(workbook,"new sdsda.xlsx")
  

//   var x=0;
//   sheet_namelist.forEach(element => {
//       var xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_namelist[x]]);

//       excelModel.insertMany(xlData,(err,data)=>{
//           if(err){
//               console.log(err);
//           }else{
//           }
//       })
//       x++;
//   });
  res.redirect('/');
});




/**
 *  @swagger
 *  /upload:
 *   post:
 *     description: 'Returns token for authorized User'
 *     operationId: Login
 *     consumes:
 *       - "application/json"
 *     requestBody:
 *        content:
 *         multipart/form-data:
 *          schema:
 *           type: object
 *           properties:
 *              fileName:
 *               type: array
 *               items:
 *                type: string
 *                format: binary
 *              
 */

app.listen(port,()=>console.log('server run at '+port));