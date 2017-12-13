function openWindowCenter(pageURL, w, h, new_name) {
	var l = eval(screen.width - w) / 2;
	var t = eval(screen.height - h) / 2;
	if (l >= (screen.width / 2)) {
		l = 0;
	}
	if (t >= (screen.height / 2)) {
		t = 0;
	}
	var _new_window = window.open("about:blank", "", "width=" + w + ",height=" + h + ",left=" + l + ",top=" + t + ",scrollbars=yes", true);
	if(document.all){
        _new_window.document.write("<head><title>«Î…‘∫Ú...</title></head><body>«Î…‘∫Ú...</body>");
    }else{
        _new_window.document.title = "«Î…‘∫Ú...";
      	_new_window.document.body.innerHTML = "«Î…‘∫Ú...";
    }
	_new_window.location.href = pageURL;
	return _new_window;
}

function checkPageNo(str) {
	if(str==null||str.length<1){
		return true;
	}else{
		if(str == 0){
			return false;
		}else{
			return /^\d+$/i.test(str);
		}
	}
}

function getXmlHttpReq() {
	var xmlHttpReq_tmp = null;
	try {
		xmlHttpReq_tmp = new XMLHttpRequest();
	}
	catch (e) {
		try {
			xmlHttpReq_tmp = new ActiveXObject("Msxml2.XMLHTTP");
		}
		catch (e1) {
			try {
				xmlHttpReq_tmp = new ActiveXObject("Microsoft.XMLHTTP");
			}
			catch (e2) {
				alert("AJAX unavailable!");
				return null;
			}
		}
	}
	return xmlHttpReq_tmp;
}

function ajaxRequest(url, callbackFunc){
	var req = getXmlHttpReq();
	req.open("post",url,true);
	req.send(null);
	
	req.onreadystatechange = function(){
   		if(req.readyState==4 && req.status == 200)
   			if(callbackFunc) {callbackFunc(req.responseText);}
   		else
   			req = null;
   	};
}

function selectAll() {
	var checkboxes = document.getElementsByName("mutilChecked");
	var isChecked = $("selectAll_check").checked;
	if (checkboxes.length) {
		for (var i = 0; i < checkboxes.length; i += 1) {
			checkboxes[i].checked = isChecked;
		}
	}
}
function selectAll_1(src_obj,OBJ_NAME) {
	var checkboxes = document.getElementsByName(OBJ_NAME);
	var isChecked = src_obj.checked;
	if (checkboxes.length) {
		for (var i = 0; i < checkboxes.length; i += 1) {
			checkboxes[i].checked = isChecked;
		}
	}
}

function how_many_selected() {
	var checkboxes = document.getElementsByName("mutilChecked");
	if (checkboxes.length) {
		var tmp1 = 0;
		for (var i = 0; i < checkboxes.length; i += 1) {
			if (checkboxes[i].checked) {
				tmp1 += 1;
			}
		}
		return tmp1;
	} else {
		return 0;
	}
}

function who_is_selected() {
	if (how_many_selected() != 1) {
		return -1;
	}
	var checkboxes = document.getElementsByName("mutilChecked");
	for (var i = 0; i < checkboxes.length; i += 1) {
		if (checkboxes[i].checked) {
			return checkboxes[i].value;
		}
	}
	return -1;
}
function who_are_selected() {
	//&mutilChecked=a&mutilChecked=b...
	if (how_many_selected() < 1) {
		return "";
	}
	var checkboxes = document.getElementsByName("mutilChecked");
	var str = "";
	for (var i = 0; i < checkboxes.length; i += 1) {
		if (checkboxes[i].checked) {
			str += ("&mutilChecked=" + checkboxes[i].value);
		}
	}
	return str;
}