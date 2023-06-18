//为jquery.serializeArray()解决radio,checkbox未选中时没有序列化的问题
//用法：取表彰序列时：$(form).hr_serialize()
$.fn.hr_serialize = function () {
	var arr_serial = this.serializeArray();
	var $radio = $('input[type=radio],input[type=checkbox]', this);
	var temp = {};
	$.each($radio, function () {
		if (!temp.hasOwnProperty(this.name)) {
			if ($("input[name='" + this.name + "']:checked").length == 0) {
				temp[this.name] = "";
				arr_serial.push({ name: this.name, value: "" });
			}
		}
	});
	//console.log(arr_serial);
	return arr_serial;						//返回数组
	//return jQuery.param(arr_serial);		//返回参数
};