/*
==轮播{对象|对象属性}==
对象属性{宽度|高度|文字大小|自动切换时间}
*/
(function($){
    dk_slideplayer2 = function(object,config){
        this.obj = object;
        this.n =0;
        this.j =0;
        var _this = this;
        var t;
        var defaults = {width:"0px",height:"0px",fontsize:"12px",bottom:"15px",time:"5000"};
        this.config = $.extend(defaults,config);
        this.count = $(this.obj + " li").size();

        if(this.config.fontsize == "14px"){
            this.size = "14px";this.height = "50px";this.right = "6px";this.bottom = "10px";
        }else{
            this.size = "12px";this.height = "30px";this.right = "6px";this.bottom = "10px";
        }

        this.factory = function(){
            //元素定位
            $(this.obj).css({position:"relative",zIndex:"0",margin:"0",padding:"0",width:this.config.width,height:this.config.height,overflow:"hidden"})
            $(this.obj).prepend("<div style='position:absolute;z-index:20;right:"+this.config.right+";bottom:"+this.config.bottom+"'></div>");
            $(this.obj + " li").css({position:"absolute",top:"0",left:"0",width:"100%",height:"100%",overflow:"hidden"}).each(function(i){
                $(_this.obj + " div").append("<a>"+(i+1)+"</a>");
            });
            $(this.obj + " img").css({border:"none",width:"100%",height:"100%"})
            this.resetclass(this.obj + " div a",0);
            //标题背景
            $(this.obj).prepend("<div class='dkTitleBg'></div>");
            //$(this.obj + " .dkTitleBg").css({position:"absolute",margin:"0",padding:"0",zIndex:"1",bottom:"0",left:"0",width:"100%",height:_this.height,background:"#000000",opacity:"0.4",overflow:"hidden"})//修改后背景颜色
            //插入标题
            $(this.obj).prepend("<div class='dkTitle'></div>");
            $(this.obj + " p").each(function(i){			
                $(this).appendTo($(_this.obj + " .dkTitle")).css({position:"absolute",margin:"0",padding:"0",zIndex:"2",bottom:"0",left:"0",width:"100%",height:_this.height,lineHeight:_this.height,textIndent:"5px",textDecoration:"none",fontSize:_this.size,color:"#ffffff",background:"none",opacity:"1",overflow:"hidden"});
                if(i!= 0){$(this).hide()}
            });
            this.slide();
            this.addhover();
            t = setInterval(this.autoplay,this.config.time);
        }
		//Download by http://www.jb51.net
        //图片渐影
        this.slide = function(){
            $(this.obj + " div a").mouseover(function(){
                _this.j = $(this).text() - 1;
                _this.n = _this.j;
                if (_this.j >= _this.count){return;}
                $(_this.obj + " li:eq(" + _this.j + ")").fadeIn("200").siblings("li").fadeOut("200");
                $(_this.obj + " .dkTitle p:eq(" + _this.j + ")").show().siblings().hide();
                _this.resetclass(_this.obj + " div a",_this.j);
            });
        }
        //滑过停止
        this.addhover = function(){
            $(this.obj).hover(function(){clearInterval(t);}, function(){t = setInterval(_this.autoplay,_this.config.time)});
        }
        //自动播放 
        this.autoplay = function(){
            _this.n = _this.n >= (_this.count - 1) ? 0 : ++_this.n;
            $(_this.obj + " div a").eq(_this.n).triggerHandler('mouseover');
        }
        //翻页函数
        this.resetclass =function(obj,i){
            $(obj).css({float:"left",width:"18px",height:"18px",lineHeight:"18px",textAlign:"center",fontWeight:"800",fontSize:"1px",color:"#fff",background:"url(Images/Ico/bn-01-01.png)",cursor:"pointer"});
            $(obj).eq(i).css({color:"#f05a28",background:"url(Images/Ico/bn-01-02.png)",textDecoration:"none"});//修改数字框背景颜色
        }
        this.factory();
    }
})(jQuery)
