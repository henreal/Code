/**
H5u Html5客户端压缩图处并上传，服务端php接收base64编码图片并保存
*/
(function($){
	$.fn.extend({
		aiiUpload:function(obj){
			if(typeof obj !="object"){
				alert('参数错误');
				return;
			}
			var imageWidth,imageHeight;
			var base64;
			var file_num=0;
			var fileInput=$(this);
			var fileInputId=fileInput.attr('id');
			var randID = Math.floor(Math.random() * 1000);	//随机ID
			createDoc('#' + fileInputId, obj.method, obj.action, obj.subText, randID, obj.formIdName, obj.imgurl);
			$('#h5u_file_'+ randID).change(function(){
				if(test(this.value)==false){
					alert('格式错误');
					return;
				}
				var objUrl = getObjectURL(this.files[0]);
				if (objUrl){		//imgBefore(objUrl,file_num);
					render(objUrl, obj.max_w, obj.max_h, file_num, randID, obj.action, obj.FaceCard);
					file_num++;
				}
			});
		}
	});
	//创建form表单
	function createDoc(objID,form_method,form_action,sub_txt,rand_id,form_id_name,form_imgurl){                
		var element=$(objID);
        //element.append('<ul class="viewList"></ul>').append('<div class="fileBox"><input type="file"  id="aii_file" /><div class="file_bg"></div></div>').append('<form id="aii_upload_form" method="'+form_method+'" action="'+form_action+'"></form>').append('<canvas id="canvas"></canvas>');
		var str = '<div class="upbox3"><canvas id="canvas_' + rand_id + '"></canvas><div class="options hide" id="h5u_options_' + rand_id + '"><a class="btn-del" id="h5u_del_' + rand_id + '">删除</a><a class="btn-big" id="h5u_big_' + rand_id + '">查看</a></div><img src="' + form_imgurl + '" id="h5u_preview_' + rand_id + '"/><div class="js-upbtn" id="h5u_upbtn_' + rand_id + '"><a href="javascript:;" class="h5u_file"><input type="file" class="up_file3" id="h5u_file_' + rand_id + '" accept="image/*">+上传图片</a><p class="a-behind">' + sub_txt + '</p></div><input type="hidden" name="' + form_id_name + '" id="h5u_form_hidden_' + rand_id + '" value="' + form_imgurl + '"><div class="status hide" id="h5u_status_' + rand_id + '"></div></div>';
		element.append(str);
                //放大
                $("#h5u_big_"+ rand_id).click(function(){
                    alert('big'+ rand_id)    
                });
                //删除1.img.src = '', 2.hidden.val = '',3.show upbtn,hide options,4. status.hide
                $("#h5u_del_"+ rand_id).click(function(){
                    $("#h5u_preview_"+ rand_id).attr('src','');
                    $("#h5u_form_hidden_"+ rand_id).val('');
                    $("#h5u_options_"+ rand_id).hide();
                    $("#h5u_upbtn_"+rand_id).show();
                    $("#h5u_status_"+ rand_id).hide();
                });
            }
            //检查文件类是否是图片格式
            function test(value){
                var regexp=new RegExp("(.JPEG|.jpeg|.JPG|.jpg|.GIF|.gif|.BMP|.bmp|.PNG|.png)$",'g');
                return regexp.test(value);
            }
            function render(src, MaximgW, MaximgH, idnum, rand_id, action, uDir) {
                console.log('rand_id',rand_id);
                var image=new Image();
                image.onload=function(){			//如果图片的尺寸小于定义的最大宽和高，则以图片原始尺寸为准
                    var canvas=document.getElementById('canvas_'+ rand_id);
                    var previewIMG = document.getElementById('h5u_preview_'+ rand_id);                   
                    //预览图片                    
                    previewIMG.src = src;
                    
                    if(image.width < MaximgW && image.height < MaximgH){
                        imageWidth = image.width;
                        imageHeight = image.height;
                    }else{
                        if(image.width>image.height){
                            imageWidth=MaximgW;
                            imageHeight=MaximgW*(image.height/image.width);
                        }else if(image.width<image.height){
                            imageHeight=MaximgH;
                            imageWidth=MaximgH*(image.width/image.height);
                        }else{
                            imageWidth=MaximgW;
                            imageHeight=MaximgH;
                        }
                    }                    
                    canvas.width=imageWidth;
                    canvas.height=imageHeight;
                    var con=canvas.getContext('2d');
                    con.clearRect(0,0,canvas.width,canvas.height);
                    con.drawImage(image,0,0,imageWidth,imageHeight);
                    base64=canvas.toDataURL('image/jpeg',0.5).substr(22);
                    //add_doc(base64,idnum);
                    //上传图片
                    upload_base(base64, action, rand_id, uDir);
                }
                image.src=src;                
            };
            //建立一個可存取到該file的url
            function getObjectURL(file) {
                var url = null ; 
                if (window.createObjectURL!=undefined) { // basic
                    url = window.createObjectURL(file) ;
                } else if (window.URL!=undefined) { // mozilla(firefox)
                    url = window.URL.createObjectURL(file) ;
                } else if (window.webkitURL!=undefined) { // webkit or chrome
                    url = window.webkitURL.createObjectURL(file) ;
                }
                return url ; 
            }
	//预览
	function imgBefore(objUrl,idnum){
		var li='<li class="view"><img src="'+objUrl+'" id="aiiImg_'+idnum+'" idnum="'+idnum+'" /><div class="close" onclick="img_remove(this);"></div></li>'
		$('.viewList').append(li);
		var img=$('#aiiImg_'+idnum);
		//预览图片居中 填满 代码
		//console.log('asdfasdfasdf');

		img.load(function(){
			var imgw=img.width(),
				imgh=img.height();
                        console.log(imgw);
                        console.log(imgh);
                    if(imgw>imgh){
                        img.css('height','100%');
                        img.css('width','auto');
                        img.css('marginLeft',-(img.width()-img.height())/2+'px');
                    }
                    else if(imgw<imgh){
                        img.css('width','100%');
                        img.css('height','auto');
                        img.css('marginTop',-(img.height()-img.width())/2+'px');
                    }
                });
            }
            //上传图片
	function upload_base(base, action, rand_id, faceDir) {		//隐藏上传控件按钮
		$('#h5u_status_'+ rand_id).html('准备上传').addClass('normal');
		$.post(action,{image:base,FaceCard:faceDir},function(json){
			console.log(json, faceDir);
			if(json.status == 1){
				$("#h5u_form_hidden_"+ rand_id).val(json.url); //表单
				$('#h5u_upbtn_'+ rand_id).hide();
				$('#h5u_options_'+ rand_id).show();                        
				$('#h5u_status_'+ rand_id).html('上传成功').addClass('success').removeClass('normal').show();
			}else{
				$('#h5u_status_'+ rand_id).html(json.info).addClass('error').removeClass('normal').show();
			}
		},'json');
	}
	function add_doc (base,idnum){
		$('#aii_upload_form').append('<input type="hidden" name="img[]" id="f_'+idnum+'" value="'+base+'"/>');
	}
})(jQuery);

function img_remove(element){
	var num = $(element).prev().attr('idnum');
	$(element).parent().remove();
	$('#f_'+num).remove();
	//console.log('asdf');
}