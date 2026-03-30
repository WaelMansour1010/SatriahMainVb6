var curX=-10;

function RGB(r,g,b)
{
    hex=new Array('0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F');
    rgb="#";
    rgb+=hex[r>>4];
    rgb+=hex[r & 15];
    rgb+=hex[g>>4];
    rgb+=hex[g & 15];
    rgb+=hex[b>>4];
    rgb+=hex[b & 15];
    return rgb;
}

var text="";

function DoEffect1(str)
{
    text=str;
    html="";
    if (curX>0)
    {
        html="<font color=white>"+str.substr(0,curX)+"</font>";
    }
    dr=(0xFF-0x20)/10;
    dg=(0xFF-0x5a)/10;
    db=(0xFF-0x87)/10;
    r=0xFF;
    g=0xFF;
    b=0xFF;
    for (i=curX;i<curX+10 && i<str.length;i++)
    {
        if (i>=0) html+="<font color="+RGB(r,g,b)+">"+str.charAt(i,1)+"</font>";
        r-=dr;
        g-=dg;
        b-=db;
    }
    if (i<str.length)
    {
        html+="<font color=#205a87>"+str.substr(i,str.length-i)+"</font>";
    }
    if (document.all['title'])
    {
        document.all['title'].innerHTML=html;
        curX++;
    }
    if (curX<str.length+11) setTimeout("DoEffect1(text)",5);
  
}

