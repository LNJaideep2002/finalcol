var xl =require("xlsx");
const generatefile=(file,reqdata,user)=>
{
    
var unifile,noofexfile,labpos;
for(i=0;i<file.length;i++)
{
    if(file[i].search("Pos")!=-1)
    {
        labpos=file[i]
    }
    if(file[i].search("No Of exper")!=-1)
    {
        noofexfile=file[i]
    }
    if(file[i].search("University")!=-1)
    {
        unifile=file[i]
    }
}
var wb1=xl.readFile("./public/"+unifile);// the marks in each cos
var ws1=wb1.Sheets["Sheet1"]; 
var data1= xl.utils.sheet_to_json(ws1);
var wb2=xl.readFile("./public/"+noofexfile); // the experiment in each cos
var ws1=wb2.Sheets["Sheet1"]; 
var data2= xl.utils.sheet_to_json(ws1);
var wbg=xl.readFile("./public/Grade.xlsx");
var wsg=wbg.Sheets["grade"]; 
var datag= xl.utils.sheet_to_json(wsg);
var wbp=xl.readFile("./public/"+labpos);
var wsp=wbp.Sheets["Sheet1"]; 
var datap= xl.utils.sheet_to_json(wsp);
// console.log(datap);

// console.log(data2);
for(i=0;i<data1.length;i++)
{
    for(j in datag[0])
    {
        if(data1[i].Grade==datag[0][j])
        data1[i]["Mark"]=j.split('m')[1];
    }
}
var exper=data1[0];
// console.log(data1.length);
data1.shift();// to remove frist element in the array
var cos={}; //no of cos

for(i in exper) // to remove & and convert it into a array
{
    exper[i]=exper[i].split("&");
}
for(i in exper) //  to remove white spaces if any are present when cos two are included
{
    for(j=0;j<exper[i].length;j++)
    {
        if(exper[i][j].split(" ").length>1)
        {
            if(exper[i][j].split(" ")[0].length>1)
            exper[i][j]=exper[i][j].split(" ")[0]
            else exper[i][j]=exper[i][j].split(" ")[1]
        }
    }
}
    // console.log(exper);
for(i=0;i<data2.length;i++)
    cos[data2[i].COs]={mark:[],exper:[],noofabovetarget:0,pernoofabovetarget:0,coatt:0,poattyorn:"y"};// mark is mark of all mark in exper and exper is the experiment which comes under respected cos
for(i in cos)
{
    for(j in exper)
        for(t=0;t<exper[j].length;t++)
        if(i==exper[j][t])
        cos[i].exper.push(j);
}
for(i in cos)
{
    for(j=0;j<data1.length;j++)
    {
        cos[i].mark[j]=0;
        for(t=0;t<cos[i].exper.length;t++)
        {
            for(m in data1[j])
            if(m==cos[i].exper[t])
            cos[i].mark[j]+=data1[j][m];
        }
        //console.log(cos[i].mark[j]*10+data1[j].Model+data1[j].Attendance);
        cos[i].mark[j]=Math.ceil((((cos[i].mark[j]*10+data1[j].Model+data1[j].Attendance)/(cos[i].exper.length+2))*0.2)+data1[j].Mark*0.8);  
        // if(j%2==1)
        // cos[i].mark[j]+=1;
    }
}

for(i in cos) // to calculate no of st above target
{
    for(j=0;j<cos[i].mark.length;j++)
    {
        if(cos[i].mark[j]>reqdata.target) //85 target
        cos[i].noofabovetarget++;
    }
}
for(i in cos) // to calculate percentage of above target and attainment level
{
    cos[i].pernoofabovetarget=Math.floor((cos[i].noofabovetarget*100)/cos[i].mark.length);
    if(cos[i].pernoofabovetarget>70)
    cos[i].coatt=3;
    else     if(cos[i].pernoofabovetarget>60)
    cos[i].coatt=2;
    else 
    cos[i].coatt=1;
}
// console.log(cos);
var psosum=0;
var psoatsum=0;
var psoattain={};
for(i in cos) // to calculate pso po coatt 
{
    for(j in datap[0])
    {
     psosum=0; 
     psoatsum=0;
     if(j!="COs")
            {
        for(t=0;t<datap.length;t++)
        {
            if(datap[t][j]!="-"){
                psoatsum+=datap[t][j]*cos[datap[t].COs].coatt;
                psosum+=datap[t][j];
            }
        }
        // console.log(psosum);
        // console.log(psoatsum);
        if(psoatsum==0)
        psoattain[j]=0;
        else
        psoattain[j]=psoatsum/psosum;
    }
    }
}

for(i in cos)
{
    if(cos[i].coatt>=2.5)
    cos[i].poattyorn="Y"
    else
    cos[i].poattyorn="N"
}
var noofexp=0;
for(i in exper)
{
    noofexp++
}
var fs=require("fs");
var html=fs.readFileSync("./headhtml.txt",{encoding:'utf8',flag:'r'});
html=html.split("<dep>").join(""+reqdata.depart);
html=html.split("<sem>").join(""+reqdata.semester);
html=html.split("<sub>").join(""+reqdata.subject);
html=html.split("<tar>").join(""+reqdata.target);
html=html.split("<br>").join(`<div><b>No. of Experiments: ${noofexp}</b><br><b>No. of COs : ${data2.length}</b><br><b>Total No. of Students :${data1.length}</b></div>`+"");
html+="<table style='border-collapse:collapse;table-layout:fixed;width:auto'><tr><td class='border'>CO's</td><td  class='border' colspan='8'>Experiment Numbers</td><td  class='border'>Total</td></tr>"
for(i=1;i<=data2.length;i++)
{
    html+="<tr>";
    html+=` <td class="border">${"CO"+i}</td>`;
    var t=0,to=0;
    for(j in exper)
    {
        for(k=0;k<exper[j].length;k++)
        {
            if(exper[j][k]==("CO"+i))
            {
                html+=` <td class="border">${j.split(".")[2]}</td>  `
                t++; 
                to++;
            }
        }
    }
    if(t!=8)
    {
        for(;t<8;t++)
        {
            html+=` <td class="border">${0}</td>`;
        }
    }
    html+=` <td class="border">${to}</td>`;
    html+="</tr>";
}
html+="</table><table style='border-collapse:collapse;table-layout:fixed;width:auto;margin:20px 0px 0px 0px'><tr>"
html+="<td class='border' rowspan='2' >S:NO</td><td class='border' rowspan='2'  >Register Number</td>";
var t=1
for(i in exper)
{
html+=`<td class='border'>${"EX:"+t}</td>`
t++
}
html+=`<td class='border' rowspan='2' >Mod</td><td class='border' rowspan='2' >Att</td><td class='border' rowspan='2' >GRA</td><td class='border' rowspan='2' >MAR</td><td class='border' colspan='${data2.length}}' >Mark obtained for Course Outcomes (COs)</td></tr><tr>`;
for(i in exper)
{
    html+=`<td class='border'>${exper[i].join("-")}</td>`
}
for(i=1;i<=data2.length;i++)
{
    html+=`<td class='border'>${"CO"+i}</td>`
}
html+="</tr>"
for(i=1;i<data1.length;i++)
{
    html+="<tr>";
    for(j in data1[i])
    html+=`<td class='border'>${data1[i][j]}</td>`;
    for(j in cos)
    {
        html+=`<td class='border'>${cos[j].mark[i-1]}</td>`;
    }
    html+="</tr>";
}

html+="<tr>";
html+=`<td class="border" style="text-align: center" colspan='${5+t}'>No. of students scored above target</td>`
for(i in cos)
{
    html+=`<td class='border'>${cos[i].noofabovetarget}</td>`;
}
html+="</tr>";
html+="<tr>";
html+=`<td class="border" style="text-align: center" colspan='${5+t}'>% of students scored above target mark</td>`
for(i in cos)
{
    html+=`<td class='border'>${cos[i].pernoofabovetarget}</td>`;
}
html+="</tr>";
html+="<tr>";
html+=`<td class="border" style="text-align: center" colspan='${5+t}'>Attainment level</td>`
for(i in cos)
{
    html+=`<td class='border'>${cos[i].coatt}</td>`;
}
html+="</tr></table><table style='border-collapse:collapse;table-layout:fixed;width:auto;margin-top: 20px;'><tr>";
var noofpos=-1;
for(j in datap[0])
{
    if(j.search("S")==-1)
    noofpos++
}
noofpos+=4;
// console.log(noofpos)
html+=`<td class='border' colspan='${noofpos}' style="text-align: center">Calculation of PO Attainment</td></tr><tr><td class="border">CO's</td><td class="border">CO Attainment</td><td class="border">Attainment</td>`
for( j in datap[0])
{
    if(j.search("S")==-1&&j!="COs")
    html+=`<td class='border'>${j}</td>`
}
html+="</tr>";
for(i=1;i<=data2.length;i++)
{
    html+="<tr>"
    html+=`<td class='border'>${"CO"+i}</td>`
    html+=`<td class='border'>${cos["CO"+i].coatt}</td>`
    html+=`<td class='border'>${cos["CO"+i].poattyorn}</td>`
    for(j in datap[i-1])
    if(j.search("S")==-1&&j!="COs")
    html+=`<td class='border'>${datap[i-1][j]}</td>`
    html+="</tr>"
}
html+="<tr><td class='border' style='text-align: center' colspan='3'>PO Attainment</td>";
for(j in psoattain)
{
    if(j.search("S")==-1&&j!="COs")
    html+=`<td class='border'>${psoattain[j]}</td>`   
}
html+="</tr></table><table style='border-collapse:collapse;table-layout:fixed;width:auto;margin-top: 20px;'><tr>";
var noofpos=0;
for(j in datap[0])
{
    if(j.search("S")!=-1)
    noofpos++
}
noofpos+=3;
html+=`<td class='border' colspan='${noofpos}' style="text-align: center">Calculation of PSO Attainment</td></tr><tr><td class="border">CO's</td><td class="border">CO Attainment</td><td class="border">Attainment</td>`
for( j in datap[0])
{
    if(j.search("S")!=-1)
    html+=`<td class='border'>${j}</td>`
}
html+="</tr>";
for(i=1;i<=data2.length;i++)
{
    html+="<tr>"
    html+=`<td class='border'>${"CO"+i}</td>`
    html+=`<td class='border'>${cos["CO"+i].coatt}</td>`
    html+=`<td class='border'>${cos["CO"+i].poattyorn}</td>`
    for(j in datap[i-1])
    if(j.search("S")!=-1)
    html+=`<td class='border'>${datap[i-1][j]}</td>`
    html+="</tr>"
}
html+="</table>"
filename=user.username+"_"+reqdata.depart+"_"+reqdata.batch+"_"+reqdata.semester+"_"+reqdata.subject+".html";
fs.writeFileSync(filename,html,"utf8");
return filename
}
exports.generatefile=generatefile;