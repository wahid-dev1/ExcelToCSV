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
app.post('/convert',(req,res)=>{
    let data=req.body.data
    let t=true
    data.map((item,i)=>{
      fs.readdir(item.csv, (err, files) => {
        if(err) {
            console.log(err)
            
              res.status(500).send({})
               return
          }
          files.forEach((file,index) => {
            if(file.split('.')[1]=="csv"){
              let source = path.join(item.csv, file);
              let destination = path.join(item.excel, file.split('.')[0]+'.xlsx');
              fs.unlink(destination,(err)=>{
                try {
                  let inboxFolders={
                  Payables:'ImportAPInvoices',
                  Vendors:'ImportAPVendors',
                  Customer:'ImportARCustomers',
                  Projects:'ImportARCustomerShipTos',
                  Invoicing:'ImportARInvoices'
                }
                  convertCsvToXlsx(source, destination);
                  const filePath =destination
                  let copy = path.join(destination,'../../')
                  copy=copy+'inbox\\'+inboxFolders[file.split('.')[0].split(' ')[0]]+
                   '\\'+file.split('.')[0]+'.xlsx'
                  
                 
                  fs.copyFile(filePath, copy, (error) => {
                    if (error) {
                        
                        } else {
                          fs.unlink(source,(err)=>{
                            console.log(err)
                         })
                          console.log('File has been moved to another folder.')
                              }
                              })
                } catch (e) {
                }
              }) 
            }
            })}
            )
            if(i==data.length-1){
              if(t){
                res.send({})
                t=false
                return
              }
            }    
    })
   
     

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

