

var request = new XMLHttpRequest();
var xmlPlayerFwd;
var PlayerFwdSprites=[];
var playerImage;
var playerImageBwd;
var PlayerSprites=[];
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
        PlayerSprites.push(PlayerSpriteObj);
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
var playerX=300;
var i=0;
var direction=1;
var keys ={
    'up':38,
    'down':40,
    'left':37,
    'right':39
};

var isKeyDown=function(keyName){
return keyDown==keys[keyName];
};


backgroundImg.src="img/background1.jpg";
var keyDown=0;
function drawPlayer(i1,direction){
    var x1=PlayerSprites[i1].x;
    var y1=PlayerSprites[i1].y;
    var width1 =PlayerSprites[i1].width;
    var height1 =PlayerSprites[i1].height;
    PlayerSprites[i1].dom=direction==1?playerImageFwd:playerImageBwd;
    canvasCtx.drawImage(PlayerSprites[i1].dom,x1,y1,width1,height1, playerX,580,width1,height1);
}

function drawBackGround(x){
    canvasCtx.clearRect(0,0,2346,769);
    canvasCtx.drawImage(backgroundImg,x,0,1346,769,0,0,2346,769);

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
    drawBackGround(x);
    drawPlayer(Math.floor(i/8),direction);
   if (isKeyDown('right')) {
       if (x<1000)
       {
           if (playerX>300)
               x++;
           else
                playerX++;

       }else{
           if (playerX<2280)
               playerX++;
       }
       i++;
       if (i>32) i=0;
       direction=1;

   };
   if (isKeyDown('left')) {
       if (x>0)
       {
           if (playerX<1800)
               x--;
           else
               playerX--;

       }else{
           if (playerX>0)
                playerX--;
       }
       i++;
       if (i>32) i=0;
       direction=2;
   };


    window.requestAnimFrame(gameLoop);
}
window.addEventListener('load', function() {
    playerImageFwd=document.createElement('img');
    playerImageFwd.src="img/playerSprite1.png";
    playerImageFwd.onload=function(){
        var id = window.requestAnimFrame(gameLoop);
    };
    playerImageBwd=document.createElement('img');
    playerImageBwd.src="img/playerSprite2.png";

});

window.addEventListener('keydown',function(ev){
    keyDown = ev.keyCode;
});
window.addEventListener('keyup',function(ev){
    console.log('down');
    keyDown = 0;
});



