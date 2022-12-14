
const windowService=require('node-windows').Service

const svc=new windowService({
    name:'AGS CSV To Excel',
    description:'Server created by AGS Advanced Inc',
    script:__dirname+"/index.js"
})

svc.on('install',()=>{
    svc.start()
})

svc.install()