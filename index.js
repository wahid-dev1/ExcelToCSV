var xlsx = require('node-xlsx');
var fs = require('fs');
const express = require('express')
const path = require('path');
const { convertCsvToXlsx } = require('@aternus/csv-to-xlsx');
const app = express()
var cors = require('cors')
app.use(cors())
const port = 7000
if (typeof localStorage === "undefined" || localStorage === null) {
  var LocalStorage = require('node-localstorage').LocalStorage;
  localStorage = new LocalStorage('./scratch');
}

app.use(express.json({extended: true}));
app.use(express.urlencoded());

app.use(express.static('./build'));
app.post('/create', (req, res) => {
       let data=localStorage.getItem('company')
       if(data){
        data=JSON.parse(data)
        data.push(req.body)
        localStorage.setItem('company',JSON.stringify(data))
       }else{
        data=[]
        data.push(req.body)
        localStorage.setItem('company',JSON.stringify(data))
       }
       res.send({
        data:JSON.parse(localStorage.getItem('company'))
       })
})
app.get('/',(req,res)=>{
  res.sendFile(path.join(process.cwd(),'./','build/index.html'))
})
app.get('/getCompanies', (req, res) => {
  
   let data=localStorage.getItem('company')
   res.send({
    data:JSON.parse(localStorage.getItem('company')||'[]')})
})
app.post('/convert', async (req, res) => {
  let data = req.body.data;
  let t = true;

  for (let i = 0; i < data.length; i++) {
    try {
      let item = data[i];
      let files = await fs.promises.readdir(item.csv);
   
      for (let j = 0; j < files.length; j++) {
        let file = files[j];
        if (file.split('.')[1] === 'csv') {
          let source = path.join(item.csv, file);
          let destination = path.join(item.excel, file.split('.')[0] + '.xlsx');
         
          let inboxFolders = {
            APInvoices: 'ImportAPInvoices',
            APVendors: 'ImportAPVendors',
            ARCustomers: 'ImportARCustomers',
            ARCustomerShipTo: 'ImportARCustomerShipTos',
            ARInvoices: 'ImportARInvoices'
          };
          convertCsvToXlsx(source, destination, {
            overwrite: true
          });
          const filePath = destination;
          console.log(filePath);
          let copy = path.join(destination, '../../');
          copy = copy + 'inbox\\' + inboxFolders[file.split('.')[0].split(' ')[0]] +
            '\\' + file.split('.')[0] + '.xlsx';
          console.log(copy);
          await fs.promises.copyFile(filePath, copy);
          await fs.promises.unlink(source);
          await fs.promises.unlink(destination);
          console.log('File has been moved to another folder.');
        }
      }
    } catch (error) {
       console.log(error.message);
      res.status(400).send({error:error.message});
      return;
    }
  }

  if (t) {
    res.send({});
    t = false;
    return;
  }
});


app.get('/remove',(req,res)=>{
  let data=JSON.parse(localStorage.getItem('company')||[])
  let index=req.query.index
 
  data.splice(index, 1)

  localStorage.setItem('company',JSON.stringify(data))
  res.send({
    data:data
  })
})
app.listen(port, () => {
  console.log(`App listening at port ${port}`)
})

