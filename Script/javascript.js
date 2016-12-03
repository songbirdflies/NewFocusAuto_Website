// JavaScript Document
// JavaScript Document
function getbak()
{
	window.location.href=document.referrer;
	return false;
}

function open_mem(id)
{
	var tabe=document.getElementById(id);
	var img=document.getElementById(id+"_pic")
	if (tabe.style.display=="none")
		{
			tabe.style.display="";	
			img.src="images/Sort_Folder_Open.gif";
		}
	else
	{
		tabe.style.display="none";
		img.src="images/Sort_Folder_Close.gif";
	}
	return false;
}