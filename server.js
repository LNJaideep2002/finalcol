var express=require('express');
var mongoose=require('mongoose');
var xl =require("xlsx");
var ejs=require("ejs");
var fs=require("fs");
var excel=require("./excel");
var excelab=require("./excelab");
var app=express();
var fileupload =require('express-fileupload');
var bodya=require("body-parser");
app.use(bodya.urlencoded({extended:true}));
app.use(bodya.json());
app.use(fileupload());
mongoose.connect('mongodb://localhost/cse',{useNewUrlParser: true, useUnifiedTopology: true});
var schema_login=mongoose.Schema({username: {type:String, unique: true },email: String,password: String});
var model_login=mongoose.model('login',schema_login);
var schema_teacher=mongoose.Schema({name: String,username: {type:String, unique: true },email: String,post: String});
var model_teacher=mongoose.model('teacher',schema_teacher);
var schema_batch=mongoose.Schema({Semester: Number,exam: Array,file: Array});
var model_batch=mongoose.model('batch',schema_batch);
var schema_subject=mongoose.Schema({Subject: Array});
var model_subject=mongoose.model('subjects',schema_subject);
app.use(express.static('public'));
app.set('view engine','ejs');
var user={
    username:"",
    password:"",
    email:"",
    name:"",
    post:"",
    generatedfile:"",
}
app.listen(6500,function()
{
    console.log("server stated");
});
app.get("/",function(req,res)
{
    res.render("login",{st:2})
});
app.post("/login",function(req,res)
{

    model_login.find({},function(err,data){
        var t=0;
        if("admin@gmail.com"==req.body.email)
        {
            if("admin"==req.body.password)
            {

                t=1;
                user.username="admin2021";
                user.password="admin";
                user.email="admin@gmail.com";
            }
            else
            {
                t=2;
        
            }
        }
        if(t==2)//email matching password not matching
        {
            res.render("login",{st:1});
        }
        if(t==1)
        {
            model_teacher.find({username:user.username},function(err,data){
                user.name=data[0].name;
                user.post=data[0].post;
                res.render("astaffd",data[0]);
            });
        }
        var t=0;
        for(i=0;i<data.length;i++)
        {
            if(data[i].email==req.body.email)
            {
                if(data[i].password==req.body.password)
                {

                    t=1;
                    user.username=data[i].username;
                    user.password=data[i].password;
                    break;
                }
                else
                {
                    t=2;
                    break;
                }
            }
            else
            {
                t=0;
            }
        }
        if(t==0)//email not found
        {
            res.render("login",{st:0});
        }
        if(t==2)//email matching password not matching
        {
            res.render("login",{st:1});
        }
        if(t==1)
        {
            model_teacher.find({username:user.username},function(err,data){
                // console.log(data[0]);
                res.render("ustaffd",data[0]);
            });
        }
    });
});
app.get("/changepass",function(req,res) {
    if(user.username=="")
    res.render("login",{st:2})
    res.render("achange",{...user,stats:0});
});
app.post("/password",function (req,res){

    if(req.body.status==0)
    {
        res.render("achange",user);
    }else{
        model_login.findOneAndUpdate({username:user.username},{password:req.body.newp},{upsert: true},function(params) {
            console.log("updated");
            if(user.username=="admin2021")
            {
                res.render("achange",{...user,stats:1}); //stats password change successfully
            }else{
                res.render("uchange",{...user,stats:1}) //stats password change successfully
            }

        })
    }
});
app.get("/addsubject",function(req,res)
{
    if(user.username=="")
    res.render("login",{st:2})
    res.render("aaddsubject",{...user,stats:0});
});
app.post("/addsubject1",function(req,res){
    if(req.body.addsub==undefined)
    {
        var file=req.files.file;
        filename="regulation"+req.body.Reg+".xlsx";
        file.mv("./public/"+filename,function(err)
        {
            var wb=xl.readFile("./public/"+filename);
            var ws=wb.Sheets["Sheet1"]; 
            var data1= xl.utils.sheet_to_json(ws);
            var batch=[];
            req.body.Reg=Number(req.body.Reg);
            console.log(data1);
            for(i=0;i<4;i++)
            {
                batch.push(req.body.Reg+i+"-"+(req.body.Reg+i+4));
            }
            var subjects=[];
            for(i=0;i<4;i++)
            {
                for(j=0;j<data1.length;j++)
                {
                    subjects.push(req.body.depart+"#"+batch[i]+"#"+"Semester "+data1[j].Semester+"#"+data1[j].Subject+"#"+data1[j].cos);
                }
            }
            var newmodel=new model_subject;
            newmodel.id=1;
            newmodel.Subject=subjects;
            model_subject.find({id:1},function(err,data)
            {
                console.log(data);
                for(i=0;i<data[0].length;i++)
                newmodel.Subject.push(data[0].Subject[i]);
                // console.log(newmodel.Subject)
                model_subject.remove({id:1},function(){
                });
                newmodel.save(function(err)
                {
                    console.log("updated");
                })
            });
            res.render("aaddsubject",{...user,stats:1});
        })
    }else{
        var subject=req.body.depart+"#"+req.body.batch+"#"+req.body.semester+"#"+req.body.sub+"#"+req.body.co;
        var newmodel=new model_subject;
        newmodel.id=1;
        newmodel.Subject=[];
        model_subject.find({id:1},function(err,data)
        {
            // console.log(data);
            newmodel.Subject=data[0].Subject;
            newmodel.Subject.push(subject);
            console.log(newmodel.Subject)
            model_subject.remove({id:1},function(){
            });
            newmodel.save(function(err)
            {
                console.log("updated");
            })
        });
        res.render("aaddsubject",{...user,stats:1});
    }
});
app.get("/addteacher",function(req,res)
{
    if(user.username=="")
    res.render("login",{st:2})
    res.render("aaddtecher",{st:2,username:user.username});
});
app.post("/aaddteacherdata",function(req,res)
{
    var newlogin=new model_login;
    var newteach=new model_teacher;
    newlogin.username=req.body.uname;
    newlogin.email=req.body.email;
    newlogin.password=req.body.pass;
    newteach.username=req.body.uname;
    newteach.email=req.body.email;
    newteach.name=req.body.name;
    newteach.post=req.body.post;
    var t=0;
    newlogin.save(function(err)
    {
        t=1;
        newteach.save(function(err){
            t=2
            if(t==0)
            {
                res.render("aaddtecher",{st:1,username:user.username});
            }
            if(t==2)
            {
                res.render("aaddtecher",{st:0,username:user.username});
            }
        });
    
    });
});
app.get("/addexam",function(req,res)
{
    if(user.username=="")
    res.render("login",{st:2})
    model_subject.find({id:1},function(err,data)
    {
    res.render("addexam",{status:0,username:user.username,subjects:data[0].Subject})
    });
});
app.post("/exam",function(req,res)
{

    var exam=req.body.depart+"#"+req.body.batch+"#"+req.body.semester+"#"+req.body.subject+"#"+req.body.exam;
    var cos="[";
    for(i in req.body)
    {
        if(i!="batch"&&i!="depart"&&i!="semester"&&i!="subject"&&i!="exam")
        {
            cos+=i;
            cos+="-";
        }
    }
    cos+="]";
    exam+=cos;
    // console.log(exam);
    model_batch.findOneAndUpdate({Semester:Number(req.body.semester.split(" ")[1])},{$push:{exam}},{upsert: true},function()
    {
        model_subject.find({id:1},function(err,data)
    {
    res.render("addexam",{status:1,username:user.username,subjects:data[0].Subject})
    });
    })
})
app.get("/uploadfile",function(req,res)
{
    if(user.username=="")
    res.render("login",{st:2})
    model_subject.find({id:1},function(err,data)
    {
        model_batch.find({},function(err,data1)
        {
            var exam=[];
            for(i=0;i<data1.length;i++)
            {
                for(j=0;j<data1[i].exam.length;j++)
                exam.push(data1[i].exam[j]);
            }
            res.render("uploadfile",{username:user.username,subjects:data[0].Subject,exam,status:0});
        });
    });
});
app.post("/uploading",function(req,res)
{
    if(req.body.pos==undefined)
    {
    var filename=req.body.depart+"#"+req.body.batch+"#"+req.body.semester+"#"+req.body.subject+"#"+req.body.exam+".xlsx";
    }else{ var filename="Pos"+"#"+req.body.depart+"#"+req.body.batch+"#"+req.body.semester+"#"+req.body.subject+".xlsx";
    }
    var file=req.files.file;
    file.mv("./public/"+filename,function(err){
        model_batch.findOneAndUpdate({Semester:Number(req.body.semester.split(" ")[1])},{$push:{file:filename}},{upsert: true},function()
        {
            model_subject.find({id:1},function(err,data)
            {
                model_batch.find({},function(err,data1)
                {
                    var exam=[];
                    for(i=0;i<data1.length;i++)
                    {
                        for(j=0;j<data1[i].exam.length;j++)
                        exam.push(data1[i].exam[j]);
                    }
                    res.render("uploadfile",{username:user.username,subjects:data[0].Subject,exam,status:1});
                });
            });
        });
    });
});
app.get("/generatefile",function(req,res)
{
    var filemiss=[]
    if(user.username=="")
    res.render("login",{st:2})
    model_subject.find({id:1},function(err,data)
    {
        res.render("generatefile",{username:user.username,subjects:data[0].Subject,filemiss,genst:0});
    });
});
app.post("/generate",function(req,res)
{
    // console.log(req.body);
    model_batch.find({},function(err,data){
        var exams1=[];
        var files=[];
        for(i=0;i<data.length;i++)
        {
            if(data[i].Semester==(req.body.semester.split(" ")[1]))
            {
                exams1=data[i].exam;
                files=data[i].file;
            }
        }
        var exams=[];
        for(i=0;i<exams1.length;i++)
        {
            var res1=exams1[i].search(req.body.depart+"#"+req.body.batch+"#"+req.body.semester+"#"+req.body.subject+"#")
            if(res1!=-1)
            {
                exams.push(exams1[i]);
            }
        }
        // console.log(exams);
        var filemiss=[];
        var fileavail=[];
        var t=0;
        for(i=0;i<exams.length;i++)
        {
            t=0;
            for(j=0;j<files.length;j++)
            {
                if(files[j]==(exams[i]+".xlsx"))
                {
                    fileavail.push(files[j]);
                    t=1;
                    break;
                }
            }
            if(t!=1)
            {
                filemiss.push(exams[i]+".xlsx")
            }
        }
       var t1=0;
        for(j=0;j<files.length;j++)
        {
            
            if(files[j]==("Pos"+"#"+req.body.depart+"#"+req.body.batch+"#"+req.body.semester+"#"+req.body.subject+".xlsx"))
           { 
               fileavail.push("Pos"+"#"+req.body.depart+"#"+req.body.batch+"#"+req.body.semester+"#"+req.body.subject+".xlsx")
               t1=1;
           }
        }
        if(t1==0)
        filemiss.push("Pos"+"#"+req.body.depart+"#"+req.body.batch+"#"+req.body.semester+"#"+req.body.subject+".xlsx")
        // console.log(fileavail);
        if(filemiss.length>=1)
        {
            // console.log(filemiss)
            model_subject.find({id:1},function(err,data)
                {
                 
                   
                    res.render("generatefile",{username:user.username,subjects:data[0].Subject,filemiss,genst:0});
                });
        }
        else if(fileavail.length==3)
        {
            model_subject.find({},function(err,data1)
            {
                var nocos=0;
                for(i=0;i<data1[0].Subject.length;i++)
                {
                    var res1=data1[0].Subject[i].search(req.body.depart+"#"+req.body.batch+"#"+req.body.semester+"#"+req.body.subject+"#");
                    if(res1!=-1)
                    {
                        nocos=Number(data1[0].Subject[i].split("#")[4]);
                    }
                }
                user.generatedfile=excelab.generatefile(fileavail,req.body,user);
                model_subject.find({id:1},function(err,data)
                {
               res.render("generatefile",{username:user.username,subjects:data[0].Subject,filemiss,genst:1});
                });
            });
        }
        else{
            model_subject.find({},function(err,data1)
            {
                var nocos=0;
                for(i=0;i<data1[0].Subject.length;i++)
                {
                    var res1=data1[0].Subject[i].search(req.body.depart+"#"+req.body.batch+"#"+req.body.semester+"#"+req.body.subject+"#");
                    if(res1!=-1)
                    {
                        nocos=Number(data1[0].Subject[i].split("#")[4]);
                    }
                }
               user.generatedfile=excel.generatefile(nocos,exams,fileavail,req.body,user);
               model_subject.find({id:1},function(err,data)
               {
               res.render("generatefile",{username:user.username,subjects:data[0].Subject,filemiss,genst:1});
               });
            })
        }
    })
});
app.get("/generatedfile",function (req,res){

    if(user.username=="")
    res.render("login",{st:2})
    res.sendFile(__dirname+"\\"+user.generatedfile);
})
