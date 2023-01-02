 //读取excel
 var wb; //读取完成的数据
 var aa = [];
 var bb=[];
 var text = [];
 var state=1;
 var rABS = false; //是否将文件读取为二进制字符串
 var picType="bar";//画图类型
 var key1,key2;//记录表头
 var u=[];
 var dd={};
 var wbsheet={};
 var pp={};
 var dataArray;
 var wbindex=[];//获取表格数据
 function typeFunction(){
     
 }
 function importExcel(obj) { //导入数据处理
     if (!obj.files) {
         return;
     }
     const IMPORTFILE_MAXSIZE = 1 * 2048; //这里可以自定义控制导入文件大小
     var suffix = obj.files[0].name.split(".")[1]
     if (suffix != 'xls' && suffix != 'xlsx') {
         alert('导入的文件格式不正确!')
         return
     }
     if (obj.files[0].size / 1024 > IMPORTFILE_MAXSIZE) {
         alert('导入的表格文件不能大于2M')
         return
     }
 
     var f = obj.files[0];
     var reader = new FileReader();
     reader.onload = e=> {
         var data = e.target.result;
         if (rABS) {
             wb = XLSX.read(btoa(fixdata(data)), { //手动转化
                 type: 'base64'
             });
         } else {
             wb = XLSX.read(data, {
                 type: 'binary'
             });
         }
         //wb.SheetNames[0]是获取Sheets中第一个Sheet的名字
         //wb.Sheets[Sheet名]获取第一个Sheet的数据
         console.log(wb.Sheets[wb.SheetNames[0]])
         dd=wb.Sheets[wb.SheetNames[0]];
         wbsheet=wb.Sheets[wb.SheetNames[0]];
         pp=JSON.stringify(XLSX.utils.sheet_to_json(dd));//原数据
         document.getElementById("databoard").innerHTML=pp;
         key1=dd.A1.v;
         key2=dd.B1.v;
        
         dd.A1.v=dd.A1.h=dd.A1.w="name";
         dd.B1.v=dd.B1.h=dd.B1.w="value";
         aa = JSON.stringify(XLSX.utils.sheet_to_json(dd));
 
         // dd.A1.v=dd.A1.h=dd.A1.w="value";
         // dd.B1.v=dd.B1.h=dd.B1.w="name";
         // bb = JSON.stringify(XLSX.utils.sheet_to_json(dd));

         var uu = eval('(' + aa + ')');
        
         u=uu;
 
         //检查文件数据格式
         var memcount=0;
         for(var x in u[0]){
             memcount++;
         }
         if(memcount!=2){
             alert("文件数据格式不符合！");
         }
         const range=XLSX.utils.decode_range(wbsheet['!ref']);
         console.log(range)//获取表格行数和列数range.e.c列数，range.e.r行数，从0开始数
         console.log(wbsheet)
         console.log(wbsheet[encodeCell(2, 3)])
         
         for(let row=0;row<=range.e.r;row++)//遍历excel表
             for(let col=0;col<=range.e.c;col++){

             }
         //渲染选择数据框
         dataArray=JSON.parse(pp);//json转为数组
         console.log(dataArray)
         var deledata=document.getElementById("deledata");
         deledata.options.length=1;
         for(var i=0;i<dataArray.length;i++){
             var dataValue="";
             for(var x in dataArray[i]){
                 dataValue+=x+":"+dataArray[i][x]+" ";
             }
             var option=new Option(dataValue,i);
             deledata.options.add(option);
         }
         
         
     
         //获取表格中为address的那列存入text中
         for (var i = 0; i < u.length; i++) {
             text.push(u[i].address);
         }
     

     };
     if (rABS) {
         reader.readAsArrayBuffer(f);
     } else {
         reader.readAsBinaryString(f);
     }
     
 }
 function encodeCell(r, c) {
     return XLSX.utils.encode_cell({ r, c });
 }
 //删除数据
 function deledata(){
     var index=document.getElementById('deledata').value;
     console.log(index)
     deleteRow(index)
     aa = JSON.stringify(XLSX.utils.sheet_to_json(wbsheet));
         u = eval('(' + aa + ')');
         console.log(u)
         //渲染选择框
     dataArray=JSON.parse(aa);//json转为数组
         console.log(dataArray)

         var deledata=document.getElementById("deledata");
         deledata.options.length=1;
         for(var i=0;i<dataArray.length;i++){
             var dataValue="";
             for(var x in dataArray[i]){
                 dataValue+=x+":"+dataArray[i][x]+" ";
             }
             var option=new Option(dataValue,i);
             deledata.options.add(option);
         }
         document.getElementById("databoard").innerHTML=aa;
 }
 //按行删除
 function deleteRow(index) {
         const range = XLSX.utils.decode_range(wbsheet['!ref']);
         for (let row = index; row < range.e.r; row++) {
             for (let col = range.s.c; col <= range.e.c; col++) {
                 wbsheet[encodeCell(row, col)] = wbsheet[encodeCell(row + 1, col)];
             }
         }
         range.e.r--;
         wbsheet['!ref'] = XLSX.utils.encode_range(range.s, range.e);
         
     }
 //画柱状图
 function drawbar(){
     picType="bar";
     echarts.dispose(document.getElementById('main'));
     myChart = echarts.init(document.getElementById('main'));
     var xDataArr=[];
     var yDataArr=[];
     console.log(u)
     for(var i=0;i<u.length;i++){//数据重新整合
             xDataArr.push(u[i].name)
             yDataArr.push(u[i].value)
     }
         myChart.setOption({
             xAxis:{
                 type:'category',
                 data:xDataArr
             },
             yAxis:{
                 type:'value'
             },
             series : [
                 {
                     type: 'bar',    // 设置图表类型为饼图
                     data:yDataArr
                 }
             ]
         })
 }
 //画折线图
 function drawline(){
     picType="line";
 
     echarts.dispose(document.getElementById('main'));
     myChart = echarts.init(document.getElementById('main'));
     myChart.hideLoading();
     var xDataArr=[];
         var yDataArr=[];
         
         for(var i=0;i<u.length;i++){
             xDataArr.push(u[i].name)
             yDataArr.push(u[i].value)
         }
             
          console.log(u)
         myChart.setOption({
         
             xAxis:{
                 type:'category',
                 data:xDataArr
             },
             yAxis:{
                 type:'value'
             },
             series : [
                 {
                     type: 'line',    // 设置图表类型为饼图
                     data:yDataArr
                 }
             ]
         })
 }
 //画饼图
 function drawpie(){
     picType="pie";
     
     echarts.dispose(document.getElementById('main'));
     myChart = echarts.init(document.getElementById('main'));
     myChart.hideLoading();
     
     myChart.setOption({
         series : [
             {
                 type: 'pie',    // 设置图表类型为饼图
                 radius: '55%',  // 饼图的半径，外半径为可视区尺寸（容器高宽中较小一项）的 55% 长度。
                 data:u
             }
         ]
         })
 }
 //画柱状标签图
 function drawbarlabel(){
    
   wbsheet=wb.Sheets[wb.SheetNames[0]];
   var app = {};
   var chartDom = document.getElementById('main');
   myChart = echarts.init(document.getElementById('main'));
   const range=XLSX.utils.decode_range(wbsheet['!ref']);
   var option;
   
   var datalabel=[];//第一行的标签
   datalabel.push(key2)
   for(let i=2;i<range.e.c+1;i++){//第一行标签
     datalabel.push(wbsheet[encodeCell(0, i)].v)
   }
//每一行的数据导入

// {
//   name: 'Forest',
//   type: 'bar',
//   barGap: 0,
//   label: labelOption,
//   emphasis: {
//     focus: 'series'
//   },
//   data: [320, 332, 301, 334, 390]
// },
       const posList = [
         'left',
         'right',
         'top',
         'bottom',
         'inside',
         'insideTop',
         'insideLeft',
         'insideRight',
         'insideBottom',
         'insideTopLeft',
         'insideTopRight',
         'insideBottomLeft',
         'insideBottomRight'
       ];
       app.configParameters = {
         rotate: {
           min: -90,
           max: 90
         },
         align: {
           options: {
             left: 'left',
             center: 'center',
             right: 'right'
           }
         },
         verticalAlign: {
           options: {
             top: 'top',
             middle: 'middle',
             bottom: 'bottom'
           }
         },
         position: {
           options: posList.reduce(function (map, pos) {
             map[pos] = pos;
             return map;
           }, {})
         },
         distance: {
           min: 0,
           max: 100
         }
       };
       app.config = {
         rotate: 90,
         align: 'left',
         verticalAlign: 'middle',
         position: 'insideBottom',
         distance: 15,
         onChange: function () {
           const labelOption = {
             rotate: app.config.rotate,
             align: app.config.align,
             verticalAlign: app.config.verticalAlign,
             position: app.config.position,
             distance: app.config.distance
           };
           myChart.setOption({
             series: [
               {
                 label: labelOption
               },
               {
                 label: labelOption
               },
               {
                 label: labelOption
               },
               {
                 label: labelOption
               }
             ]
           });
         }
       };
       const labelOption = {
         show: true,
         position: app.config.position,
         distance: app.config.distance,
         align: app.config.align,
         verticalAlign: app.config.verticalAlign,
         rotate: app.config.rotate,
         formatter: '{c}  {name|{a}}',
         fontSize: 15,
         rich: {
           name: {}
         }
       };
       //每一行的数据
       var partData=new Array();
           var rawdata=new Array();
           var databox=[][range.e.c];
       function part(name,data){
               this.name=name;
               this.data=data;
               this.type='bar';
               this.emphasis={
                   focus: 'series'
               }
               this.label=labelOption
         
           }
           
           for(let i=0;i<datalabel.length;i++){
             
               for(let j=1;j<range.e.c+1;j++){
                   rawdata.push(wbsheet[encodeCell(i+1, j)].v)
               }
               // partData.push(new part(datalabel[i],databox[i+1]));
               for(let j=1;j<range.e.c+1;j++){
                   rawdata.pop()
               }
           }
           // console.log(partData)

       option = {
         tooltip: {
           trigger: 'axis',
           axisPointer: {
             type: 'shadow'
           }
         },
         legend: {
           data: datalabel
         },
         toolbox: {
           show: true,
           orient: 'vertical',
           left: 'right',
           top: 'center',
           feature: {
             mark: { show: true },
             dataView: { show: true, readOnly: false },
             magicType: { show: true, type: ['line', 'bar', 'stack'] },
             restore: { show: true },
             saveAsImage: { show: true }
           }
         },
         xAxis: [
           {
             type: 'category',
             axisTick: { show: false },
             data: ['2012', '2013', '2014', '2015', '2016']
           }
         ],
         yAxis: [
           {
             type: 'value'
           }
         ],
         series: [
           {
             name: 'Forest',
             type: 'bar',
             barGap: 0,
             label: labelOption,
             emphasis: {
               focus: 'series'
             },
             data: [320, 332, 301, 334, 390]
           },
           {
             name: 'Steppe',
             type: 'bar',
             label: labelOption,
             emphasis: {
               focus: 'series'
             },
             data: [220, 182, 191, 234, 290]
           },
           {
             name: 'Desert',
             type: 'bar',
             label: labelOption,
             emphasis: {
               focus: 'series'
             },
             data: [150, 232, 201, 154, 190]
           },
           {
             name: 'Wetland',
             type: 'bar',
             label: labelOption,
             emphasis: {
               focus: 'series'
             },
             data: [98, 77, 101, 99, 40]
           }
         ]
       };

       option && myChart.setOption(option);
     
};
 
 //交换变量
 function switchvar(){
     if(state==1){
             u = eval('(' + bb + ')');
             state=2;
         }
         else if(state==2){
             u = eval('(' + aa + ')');
             state=1;
         }
     if(picType=="bar")
         drawbar();
     else if(picType=="line")
         drawline();
     else if(picType=="pie")
         drawpie();
 }