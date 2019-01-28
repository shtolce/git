var request = new XMLHttpRequest();
//var requestNPC = new XMLHttpRequest();
var xmlPlayerFwd;
var xmlNPCPlayer=new Map();
var PlayerFwdSprites=[];
var playerImage;
var NPCImage;
var NPCImageBWD=new Map();
var NPCImageFWD=new Map();
var playerImageBwd;
var PlayerSprites=[];
var NPCSprites=new Map();
var npcDialogueSprite;
const sceneMoveTresholdRight=300;
const sceneMoveTresholdLeft=550;
const totalSceneWidth=2346;
const visibleSceneWidth=1346;
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
                height:el.attributes['h'].value,
                dialogue:''
                };
        PlayerSprites.push(PlayerSpriteObj);
        }
    });
}

/* этот аналог правильнее. нет ненужных элементов
Array.from(document.querySelector('animation').children)
    .map((el) => Array.from(el.attributes)
        .reduce((acc, attribute) => {
            acc[attribute.name] = attribute.value
            return acc
        },

            {}))
*/


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
function reqReadyStateChangeNPC() {
    requestNPC=this;
    if (requestNPC.readyState == 4) {
        var status = requestNPC.status;
        if (status == 200) {
            xmlNPCPlayer.set(requestNPC.instance,requestNPC.responseXML);
            writeNPCObj(requestNPC.instance);
        } else {
        }
    }
}
function InitNPC(instance){
    NPCImageFWD.set(instance,document.createElement('img'));
    NPCImageFWD.get(instance).src='img/'+instance+'SpriteFWD.png';
    NPCImageBWD.set(instance,document.createElement('img'));
    NPCImageBWD.get(instance).src='img/'+instance+'SpriteBWD.png';
    var requestNPC = new XMLHttpRequest();
    requestNPC.open("GET", "img/"+instance+".xml");
    requestNPC.instance=instance;
    requestNPC.onreadystatechange = reqReadyStateChangeNPC;
    requestNPC.send();
}

window.addEventListener('load', function() {
    playerImageFwd=document.createElement('img');
    playerImageFwd.src="img/playerSprite1.png";
    playerImageFwd.onload=function(){
        var id = window.requestAnimFrame(gameLoop);
    };
    playerImageBwd=document.createElement('img');
    playerImageBwd.src="img/playerSprite2.png";
    npcDialogueSprite=document.createElement('img');
    npcDialogueSprite.src="img/DialogueSprite.png";
    request.open("GET", "img/playerFWD.xml");
    request.onreadystatechange = reqReadyStateChange;
    request.send();
    InitNPC('Seller');
    InitNPC('PoliceMan');

});
function writeNPCObj(npcInstance){
    var xmlNPCPlayer1=xmlNPCPlayer.get(npcInstance);
    var xmlAnim = xmlNPCPlayer1.getElementsByTagName("animation")[0];
    var xmlCur=xmlAnim.childNodes;
    NPCSprites.set(npcInstance,[]);
    xmlCur.forEach((el,i,xmlCur)=> {
        if (el.nodeName=="cut"){
            var NPCSpriteObj={
                domFWD:NPCImageFWD.get(npcInstance),
                domBWD:NPCImageBWD.get(npcInstance),
                dom:NPCImage,
                width:el.attributes['w'].value,
                x:el.attributes['x'].value,
                y:el.attributes['y'].value,
                height:el.attributes['h'].value,
                npcXPos:700,
                dialogue:''


            };
            NPCSprites.get(npcInstance).push(NPCSpriteObj);
        }
    });
}
var canvas=document.getElementById("mainCanvas");
var canvasCtx = canvas.getContext("2d");
var backgroundImg = new Image(totalSceneWidth,769);
var x=0;
var playerX=300;
var i=0;
var iNpc=0;
var iNpc1=0;

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
function drawNPC(npcNumber,i1,instance,direction){
    var NPCSprites1=NPCSprites.get(instance);

    var x1=NPCSprites1[i1].x;
    var y1=NPCSprites1[i1].y;
    var width1 =NPCSprites1[i1].width;
    var height1 =NPCSprites1[i1].height;
    var npcX=NPCSprites1[i1].npcXPos-(x*totalSceneWidth/visibleSceneWidth);
    if (direction==1)
        canvasCtx.drawImage(NPCSprites1[i1].domFWD,x1,y1,width1,height1, npcX,560,width1,height1);
    else
        canvasCtx.drawImage(NPCSprites1[i1].domBWD,x1,y1,width1,height1, npcX,560,width1,height1);

    if (NPCSprites1[i1].dialogue!=''){
        canvasCtx.drawImage(npcDialogueSprite,0,0,npcDialogueSprite.width,npcDialogueSprite.height, npcX,460,
            npcDialogueSprite.width,npcDialogueSprite.height);
        canvasCtx.fillStyle = "#00F";
        canvasCtx.strokeStyle = "#F00";
        canvasCtx.font = "bold italic 15pt Arial";
        canvasCtx.fillText(NPCSprites1[i1].dialogue, npcX+50, 500);

    }

}
function drawBackGround(x){
    canvasCtx.clearRect(0,0,totalSceneWidth,769);
    canvasCtx.drawImage(backgroundImg,x,0,visibleSceneWidth,769,0,0,totalSceneWidth,769);
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
function movePlayerFWD(){
    if (x<1000)
    {
        if (playerX>sceneMoveTresholdRight){
            x++;
        }
        else{
            playerX++;
        }
    }
    else{
        if (playerX<totalSceneWidth-60){
            playerX++;
        }
    }
    i++;
    if (i>32){
        i=0;
    }
    direction=1;
}
function movePlayerBWD(){
    if (x>0)
    {
        if (playerX<totalSceneWidth-sceneMoveTresholdRight){
            x--;
        }
        else{
            playerX--;
        }

    }else{
        if (playerX>0){
            playerX--;
        }
    }
    i++;
    if (i>32) {
        i=0;
    };
    direction=2;
}

function gameLoop() {
    var NPCPolice=NPCSprites.get('PoliceMan');
        NPCPolice.forEach((el,i,NPCPolice)=>{el.npcXPos=1300;el.dialogue='Fuck'});

   drawBackGround(x);
   drawNPC(1,Math.floor(iNpc/64),'Seller',direction);
   drawNPC(1,Math.floor(iNpc1/8),'PoliceMan',direction);
   drawPlayer(Math.floor(i/8),direction);
    iNpc++;
    if (iNpc>384) {
        iNpc=0;
    };
    iNpc1++;
    if (iNpc1>32) {
        iNpc1=0;
    };//

   if (isKeyDown('right')) {
       movePlayerFWD();
   };
   if (isKeyDown('left')) {
        movePlayerBWD();
   };
    window.requestAnimFrame(gameLoop);
}
window.addEventListener('keydown',function(ev){
    keyDown = ev.keyCode;
});
window.addEventListener('keyup',function(ev){
    console.log('down');
    keyDown = 0;
});



