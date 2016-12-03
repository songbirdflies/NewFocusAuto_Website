/*按比例生成缩略图*/
function DrawImage(ImgD,W,H){ 
  var flag=false; 
  var image=new Image(); 
  image.src=ImgD.src; 
  if(image.width>0 && image.height>0){ 
    flag=true; 
    if(image.width/image.height>= W/H){ 
      if(image.width>W){
        ImgD.width=W; 
        ImgD.height=(image.height*H)/image.width; 
      }
	  else{ 
        ImgD.width=image.width;
        ImgD.height=image.height; 
      } 
      ImgD.alt= ""; 
    } 
    else{ 
      if(image.height>H){
        ImgD.height=H; 
        ImgD.width=(image.width*W)/image.height; 
      }
	  else{ 
        ImgD.width=image.width;
        ImgD.height=image.height; 
      } 
      ImgD.alt=""; 
    } 
  }
}


<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->


function openwin(selID) {
var obj=document.getElementById("selID");
var url=obj.options[obj.selectedIndex].value;
window.open (url);
//bj.options[obj.selectedIndex].selected=true;
}

<!--
function show(div){ 
document.getElementById(div).style.display='block';  //显示层 
} 

function hide(div){ 
document.getElementById(div).style.display='none';  //隐藏层 
} 
//-->
/*********************************************
功能:    通用DIV切换函数
参数:    divID --当前DIV的ID号；divName --要改变的这一组DIV的命名前缀；zDivCount --这一组DIV的个数-1；On 切定图标显示css；Off 离开图标显示Css
*********************************************/
function ChangeDiv(divId,divTName,divCName,zDivCount,On,Off) 
{ 
for(i=0;i<=zDivCount;i++)
{
document.getElementById(divTName+i).className=Off; 
document.getElementById(divCName+i).style.display="none"; 
}
document.getElementById(divTName+divId).className=On; 
document.getElementById(divCName+divId).style.display="block"; 
}


function OpenScript(url,width,height)
{
  var win = window.open(url,"SelectToSort",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
}


function ScrollDivUpLeft(){document.getElementById('itemsleft').scrollTop -= 1;}
function ScrollDivDownLeft(){document.getElementById('itemsleft').scrollTop += 1;}

function ScrollDivUpRight(){document.getElementById('itemsright').scrollTop -= 1;}
function ScrollDivDownRight(){document.getElementById('itemsright').scrollTop += 1;}