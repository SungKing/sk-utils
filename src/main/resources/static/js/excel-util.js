
/**
 * import jquery-1.10.2.js
 * import xlsx.full.min.js
 * eg:
 *
 * function down_legal_file(e) {
    var that = $(e);
    var name = that.attr('filename')
    var data = [{"name":"名称","type":"类型","bank":"银行","account":"帐号","parentName":"上级"}];
    var header = [];
    for(var tmp_name in data[0]){
        header.push(tmp_name);
    }

    var trs = that.parent().find("tbody").find("tr");
    for(var i = 0;i<trs.length;i++){
        var tds = $(trs.get(i)).find("td");

        var temp = {"name":$(tds.get(0)).text(),
            "type":$(tds.get(1)).text(),
            "bank":$(tds.get(2)).text(),
            "account":$(tds.get(3)).text(),
            "parentName":$(tds.get(4)).text()}
        data.push(temp);
    }

    var date = new Date();
    var down_name = name+date.getFullYear()+"" + (date.getMonth()+1) + ""+date.getDay()+ date.getHours()+date.getMinutes()+date.getSeconds()+".xlsx";
    downExcel(header,data,down_name,'aaa')
}
 *
 *
 * 生成excel
 * @param header 引用属性
 * @param data   数据
 * @param down_name 下载名称
 */
function downExcel(header,data,down_name,download_id) {

    var workBook = XLSX.utils.book_new();
    var workSheet = XLSX.utils.json_to_sheet(data,{header:header, skipHeader:true});
    XLSX.utils.book_append_sheet(workBook, workSheet, "sheet1");
    //生成excel
    var s2ab = function (s) { // 字符串转字符流
        var buf = new ArrayBuffer(s.length)
        var view = new Uint8Array(buf);
        for (var i = 0; i !== s.length; ++i) {
            view[i] = s.charCodeAt(i) & 0xFF
        }
        return buf
    };
    // 创建二进制对象写入转换好的字节流
    var tmpDown =  new Blob([s2ab(XLSX.write(workBook,{bookType: 'xlsx', bookSST: false, type: 'binary'} ))], {type: ''});
    // var a = document.createElement('a');
    // // 利用URL.createObjectURL()方法为a元素生成blob URL
    // a.href = URL.createObjectURL(tmpDown)  // 创建对象超链接
    //
    // a.download = down_name;
    //解决火狐问题
    $('#'+download_id).attr("href",URL.createObjectURL(tmpDown));
    $('#'+download_id).attr("download",down_name);

    document.getElementById(download_id).click();
    $('#'+download_id).removeAttr('href');
    $('#'+download_id).removeAttr('download');
}