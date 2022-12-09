var xlsx = require('node-xlsx');
var fs = require('fs');
const express = require('express')
const session = require("express-session")
const flash = require("connect-flash")
const app = express()
const port = 7000
const store=require('store')
app.use(express.json({extended: true}));
app.use(express.urlencoded());
app.use(session({
    secret: "Hi",
    resave: false,
    saveUninitialized: false,
    cookie: {
        maxAge:1000*5
    }
    //* maxAge: verilen cookienin ne kadar zaman sonra kendisini ihma etmesini söylüyor
    //* saniye cinsinden verdik 
}))
app.use(flash());
app.set('view engine', 'ejs')
app.use(express.static(__dirname + '/public'));

app.post('/setdir', (req, res) => {
        store.set("excel",req.body.excel_dir)
        store.set("csv",req.body.csv_dir)
        res.redirect("/");
})
app.get('/', (req, res) => {
   
    res.render('pages/index.ejs',{
        excel:store.get('excel')||"",
        csv:store.get('csv')||"",
        message:req.flash('message')
    })
})
app.post('/convert',(req,res)=>{
    fs.readdir(store.get('excel'), (err, files) => {
        if(err) {
            console.log(err)
              req.flash("message","Error")
              res.redirect("/"); 
               return
          }
        if(files.length==0){
            req.flash("message","Error")
            res.redirect("/"); 
             return 
        }  
        files.forEach((file,index) => {
            if(file.split('.')[1]=="xlsx"){
              
              var obj = xlsx.parse(store.get('excel')+"/"+file); // parses a file
              var rows = [];
              var writeStr = "";
              
              //looping through all sheets
              for(var i = 0; i < obj.length; i++)
              {
                  var sheet =obj[i];
                  //loop through all rows in the sheet
                  for(var j = 0; j < sheet['data'].length; j++)
                  {
                          //add the row to the rows array
                          rows.push(sheet['data'][j]);
                  }
              }
              
              //creates the csv string to write it to a file
              for(var i = 0; i < rows.length; i++)
              {
                  writeStr += rows[i].join(",") + "\n";
              }
              
              //writes to a file, but you will presumably send the csv as a      
              //response instead
              fs.writeFile(store.get('csv') +"/"+file.split(".")[0]+".csv", writeStr, function(err) {
                  if(err) {
                    console.log(err)
                      req.flash("message","Error")
                      res.redirect("/"); 
                       return
                  }
                 if(index==files.length-1){
                    req.flash("message","Success")
                    res.redirect("/");              
                 }    
                  
              });
            }
          });
      });
      
  });

app.listen(port, () => {
  console.log(`App listening at port ${port}`)
})

