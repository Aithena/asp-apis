### 由ASP编写数据接口，返回json/jsonp数据
将文件下载后即可使用，方便本地需要使用到数据的测试项目来使用。

### 设置数据库
<pre>
DBPath = Server.MapPath("./#data.mdb")
</pre>

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

+ dataType: 'jsonp'
+ jsonpCallback: 'callback'


### 部署服务器
+ ASP - 编译 - 调试属性 - 将错误发送到浏览器
+ ASP - 行为 - 启用父路径
+ 应用池 - 高级设置 - 启动32位