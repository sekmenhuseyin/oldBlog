function delText(SearchText){if (document.getElementById(SearchText).value==SearchText){document.getElementById(SearchText).value='';document.getElementById(SearchText).style.color='Black';};}
function writeText(SearchText){if (document.getElementById(SearchText).value==''){document.getElementById(SearchText).value=SearchText;document.getElementById(SearchText).style.color='Gray';};}
function RefreshImage(valImageId){var objImage=document.images[valImageId];if(objImage==undefined){return;}var now=new Date();objImage.src=objImage.src.split('?')[0]+'?x='+now.toUTCString();}