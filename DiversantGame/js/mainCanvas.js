

var request = new XMLHttpRequest();
var xmlPlayerFwd;
var PlayerFwdSprites=[];
var playerImage;

function writePlayerObj(){
    var xmlAnim = xmlPlayerFwd.getElementsByTagName("animation")[0];
    var xmlCur=xmlAnim.childNodes;
    xmlCur.forEach((el,i,xmlCur)=> {
        if (el.nodeName=="cut"){
            var PlayerSpriteObj={
                dom:playerImage,
                width:el.attributes['w'].value,
                x:el.attributes['x'].value,
                y:el.attributes['y'].value,
                height:el.attributes['h'].value
        };
        PlayerFwdSprites.push(PlayerSpriteObj);
        }

    });
}

function reqReadyStateChange() {
    if (request.readyState == 4) {
        var status = request.status;
        if (status == 200) {
           xmlPlayerFwd=request.responseXML;
            writePlayerObj();
        } else {
        }
    }
}

request.open("GET", "img/playerFWD.xml");
request.onreadystatechange = reqReadyStateChange;
request.send();

var canvas=document.getElementById("mainCanvas");
var canvasCtx = canvas.getContext("2d");
var backgroundImg = new Image(2346,769);
var x=0;
var i=0;
backgroundImg.src="img/background1.jpg";

function drawPlayer(){
    i++;
    if (i>4) i=0;
    var x1=PlayerFwdSprites[i].x;
    var y1=PlayerFwdSprites[i].y;
    var width1 =PlayerFwdSprites[i].width;
    var height1 =PlayerFwdSprites[i].height;
   canvasCtx.drawImage(playerImage,x1,y1,width1,height1, 0,580,width1,height1);
}

function drawBackGround(x){
    canvasCtx.clearRect(0,0,2346,769);
    canvasCtx.drawImage(backgroundImg,x,0,1346+x,769,0,0,2346,769);
}

window.requestAnimFrame = (function(){
    return  window.requestAnimationFrame       ||
            window.webkitRequestAnimationFrame ||
            window.mozRequestAnimationFrame    ||
            window.oRequestAnimationFrame      ||
            window.msRequestAnimationFrame     ||
        function(callback){
            window.setTimeout(callback, 1000 / 60);
        };
})();


function gameLoop() {
   x++;
    drawBackGround(x);
    drawPlayer();
   window.requestAnimFrame(gameLoop);
}
window.addEventListener('load', function() {
    playerImage=document.createElement('img');
    playerImage.src="img/playerSprite1.png";
    playerImage.onload=function(){
        var id = window.requestAnimFrame(gameLoop);
    };

});


