var chartArr=[point1,point2,point3,point4];
var canv = document.getElementById('myCanvas');
var contG=canv.getContext('2d');
var contG1=canv.getContext('2d');
const width = canv.clientWidth;
const height = canv.clientHeight;
contG.moveTo(0,height);

function line(x1,y1){
    var x = x1;
    var y =height -y1;
    contG.lineTo(x,y);
    contG.stroke();
}
function getDaysAmount(datB,datE)
{
    return Math.ceil(Math.abs(datB.getTime() - datE.getTime()) / (1000 * 3600 * 24));
}
function dayToX(dat,datB,datE){
return Math.ceil(width/getDaysAmount(datB,datE))*getDaysAmount(datB,dat);
}
function getX(dat){
    return dayToX(dat,point1.x,point8.x);
}

var point1={
    x:new Date(2019, 1, 1),
    y:13
}
var point2={
    x:new Date(2019, 1, 2),
    y:43
}
var point3={
    x:new Date(2019, 1, 4),
    y:33
}
var point4={
    x:new Date(2019, 1, 12),
    y:2
}
var point5={
    x:new Date(2019, 1, 22),
    y:42
}
var point6={
    x:new Date(2019, 2, 1),
    y:2
}
var point7={
    x:new Date(2019, 2, 5),
    y:32
}
var point8={
    x:new Date(2019, 2, 12),
    y:52
}


//alert(dayToX(point1.x,point4.x));
contG.save();
contG.scale(1,0.5);
line(getX(point1.x),point1.y);
line(getX(point2.x),point2.y);
line(getX(point3.x),point3.y);
line(getX(point4.x),point4.y);
line(getX(point5.x),point5.y);
line(getX(point6.x),point6.y);
line(getX(point7.x),point7.y);
line(getX(point8.x),point8.y);
cnvobj1=document.getElementById("canv2");
imgd = contG.getImageData(0,0,500,350);
ctx1=cnvobj1.getContext("2d");
ctx1.putImageData(imgd, 0,0);
contG.restore();
contG.clear();
contG.moveTo(0,height);
line(getX(point1.x),point1.y);
line(getX(point2.x),point2.y);
line(getX(point3.x),point3.y);
line(getX(point4.x),point4.y);
line(getX(point5.x),point5.y);
line(getX(point6.x),point6.y);
line(getX(point7.x),point7.y);
line(getX(point8.x),point8.y);



