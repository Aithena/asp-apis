# 由asp编写数据接口，返回json/jsonp数据
将文件下载后即可使用，方便本地需要使用到数据的测试项目来使用。

### javascript获取json
<pre>
$.ajax({
  url: '/api/e.asp',
  data: {
    action: 'list'
  },
  type: 'post',
  dataType: 'json',
  success: function(res){
    console.log(res.code);
  },
  error: function(){
    console.log('ERROR: 出现错误');
  }
})
</pre>

### javascript获取jsonp
<pre>
$.ajax({
  url: '/api/e.asp',
  data: {
    action: 'list'
  },
  type: 'post',
  dataType: 'jsonp',
  jsonpCallback: 'callback',
  success: function(res){
    console.log(res.code);
  },
  error: function(){
    console.log('ERROR: 出现错误');
  }
})
</pre>
