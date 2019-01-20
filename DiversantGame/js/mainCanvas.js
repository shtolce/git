var canvas=document.getElementById("mainCanvas");
var canvasCtx = canvas.getContext("2d");
var backgroundImg = new Image(2346,769);
var x=0;

backgroundImg.src="img/background1.jpg";

window.addEventListener('load', function() {
    drawBackGround();

});

function drawBackGround(){
   // canvasCtx.drawImage(backgroundImg,0,0,1346,769,0,0,2346,769);
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
    canvasCtx.clearRect(0,0,2346,769);
    canvasCtx.drawImage(backgroundImg,x,0,1346+x,769,0,0,2346,769);

    window.requestAnimFrame(gameLoop);
}
var id = window.requestAnimFrame(gameLoop);