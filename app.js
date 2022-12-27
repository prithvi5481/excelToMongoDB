var express = require('express');
var mongoose = require('mongoose');
var bodyParser = require('body-parser');
var path = require('path');
var XLSX = require('xlsx');
var multer = require('multer');


//multer
var storage = multer.diskStorage({
    destination: function(req,file,cb){
        cb(null,'./public/uploads')
    },
    filename: function(req,file,cb){
        cb(null,file.originalname)
    }
});

const upload = multer({ storage: storage });

//connect to DB
mongoose.set('strictQuery', false);
mongoose.connect("mongodb://0.0.0.0:27017/excel",{useNewUrlParser:true})
.then(()=>{console.log('connected to database')})
.catch((err)=>console.log('error',err));

var app = express();

//static folder path
app.use(express.static(path.resolve(__dirname,'public')));
//set the ejs template engine
app.set('view engine','ejs');
app.use(bodyParser.urlencoded({extended:false}));


//collection schema
var excelSchema = new mongoose.Schema({
    name:String,
    email:String,
    mobile:String,
    dob:String,
    experince:String,
    title:String,
    location:String,
    address:String,
    employer:String,
    destination:String
});

var excelModel = mongoose.model('excelModel',excelSchema);

app.get('/',(req,res)=>{
    res.render('home.ejs');
});

app.post('/',upload.single('excel'),(req,res)=>{
    var workbook = XLSX.readFile(req.file.path);
    var sheet_namelist = workbook.SheetNames;
    var x=0;
    sheet_namelist.forEach(element => {
        var xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_namelist[x]]);
        excelModel.insertMany(xlData,(err,data)=>{
            if(err){
                console.log(err);
            }
            else{
                console.log(xlData);
            }
        })
        x++;
    });
    res.redirect('/');
});

var port = process.env.PORT || 5000;
app.listen(port,()=>{
    console.log(`Server is running on port ${port}`);
});
