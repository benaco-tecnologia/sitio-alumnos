
//Script created by Jim Young (www.requestcode.com)
//Submitted to JavaScript Kit (http://javascriptkit.com)
//Visit http://javascriptkit.com for this script

//Set the tool tip message you want for each link here.
     var tip=new Array
           tip[0]='Visit Dynamic Drive<br> for DHTML Scripts!'
           tip[1]='Visit JavaScript Kit for Great<br> Scripts, Tutorials and Forums!'
           tip[2]='Visit Request Code for FREE JavaScripts!'
           tip[3]='Click here for some excellent<br>Java applets and tutorials'
           
     function showtip(current,e,num)
        {
         if (document.layers) // Netscape 4.0+
            {
             theString="<DIV CLASS='ttip'>"+num+"</DIV>"
             document.tooltip.document.write(theString)
             document.tooltip.document.close()
             document.tooltip.left=e.pageX+14
             document.tooltip.top=e.pageY+2
             document.tooltip.visibility="show"
            }
         else
           {
            if(document.getElementById) // Netscape 6.0+ and Internet Explorer 5.0+
              {
               elm=document.getElementById("tooltip")
               elml=current
               elm.innerHTML="<b><font size='1' face='Verdana, Arial, Helvetica, sans-serif'>"+num+"</font></b>"
               elm.style.height=elml.style.height
               elm.style.top=parseInt(elml.offsetTop+elml.offsetHeight)
               elm.style.left=parseInt(elml.offsetLeft+elml.offsetWidth+10)
               elm.style.visibility = "visible"
              }
           }
        }
function hidetip(){
if (document.layers) // Netscape 4.0+
   {
    document.tooltip.visibility="hidden"
   }
else
  {
   if(document.getElementById) // Netscape 6.0+ and Internet Explorer 5.0+
     {
      elm.style.visibility="hidden"
     }
  } 
}
