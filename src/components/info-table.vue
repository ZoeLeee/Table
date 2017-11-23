<template>
  <el-row :style="{ fontSize: fontSize + 'px' }"> 
    <el-col :span="18">
      <div>
        <button @click="editHeader($event)" class="btn btn-primary">编辑</button>
        <button @click="addCol" class="btn btn-default">新增</button>
        <button @click="onexport" id="export-table" class="btn btn-default">导出</button>
      </div>
      <div id="out-table" v-if="headerData[value]">
        <ul class="list-unstyled list-inline pageItem" >
          <li v-for="(item,index) in pageHead" :key="item.id">
            <button type="button" class="close" aria-label="Close" v-if="item.isEdit" @click="delItem(index)">
              <span aria-hidden="true">&times;</span>
            </button> 
            <span @dblclick="editItem(index)" v-if="!item.isEdit">{{item.name}}</span>
            <input type="text" v-model="item.name" v-if="item.isEdit" @keyup.enter="editItem(index)"> 
            <input type="text" v-model="pageHeadContent[index].name">    
          </li>
          <li>
            <button class="btn btn-primary" @click="addHeadItem">+</button>
          </li>
        </ul>
        <!-- <table>
          <tr style="word-wrap:break-word;word-break:break-all">
            <td style="word-wrap:break-word;word-break:break-all;width:33%;white-space:nowrap" v-for="(item,index) in pageHead" :key="item.id">
              <span @dblclick="editItem(index)" v-if="!item.isEdit">{{item.name}}</span>
              <input type="text" v-model="item.name" v-if="item.isEdit" @keyup.enter="editItem(index)"> 
              <input type="text">
            </td>
            <td style="word-wrap:break-word;word-break:break-all;width:33%;white-space:nowrap">
              <button class="btn btn-primary" @click="addHeadItem">+</button>
            </td>
          </tr>
        </table> -->
        <table class="table table-bordered" ref="content">
          <thead ref="tHead">
            <tr>
              <template v-for="(item,key) in headerData[value]">        
                <th   :key="item.id" draggable="true">
                  <span v-if="isEdit">{{item}}</span> 
                  <input type="text" v-model="headerData[value][key]" v-if="!isEdit">
                  <button type="button" class="close" aria-label="Close" @click="delCol(key)" v-if="!isEdit">
                    <span aria-hidden="true">&times;</span>
                  </button> 
                </th>
              </template>
            </tr>
          </thead>
          <tbody ref="tbody">
            <tr v-for="data in allData[value]" :key="data.id">
              <td v-for="item in data" :key="item.id">{{item}}</td>
            </tr>
          </tbody>
          <tfoot v-if="allData[value]">
            <tr>
              <td>
                合计
              </td>
              <td :colspan="Object.keys(headerData[value]).length-1">
                <span>¥{{totalPrice}}</span>
              </td>
            </tr> 
          </tfoot> 
        </table> 
        <table>
          <tr>
            <td align="right">
              合计：<span>¥{{totalPrice}}</span>
            </td>
          </tr>    
        </table>
      </div>
    </el-col>
    <el-col :span="6">
      <el-card class="box-card" body-style="{ padding: 30px }">
        <div slot="header" class="clearfix">
          <el-select v-model="value" placeholder="请选择" @change="onSelect" ref="selectedItem">
            <el-option ref="selectedItem"
              v-for="item in options"
              :key="item.index"
              :label="item.label"
              :value="item.value">
            </el-option>
          </el-select>
        </div>
        <div>
          <el-collapse v-model="activeNames">
            <el-collapse-item title="格式设置" name="1" class="setting">
              <form action="">
                <label for="">字体大小:</label>
                <el-input-number v-model="fontSize" :min="0" size="small" label="字体大小"></el-input-number>
              </form>
            </el-collapse-item>
            <el-collapse-item  name="2">
              <template slot="title">
                {{settingPanelTitle}}
              </template>
            </el-collapse-item>
          </el-collapse>
          <el-button type="primary">立即添加</el-button>
          <el-button>取消</el-button>
        </div>
      </el-card>
    </el-col> 
  </el-row>
</template>

<script>
  import XLSX from 'xlsx';
  import XLSX_SAVE from  'file-saver';

  export default {
    data() {
      return {
        fontSize:14, //字体大小
        //切换手风琴列表
        activeNames: ['1'],
        // 设置面板标题
        settingPanelTitle:"",
        col:1,
        // 是否编辑
        isEdit:true,
        //总额
        totalPrice:0,
        // 表头数据
        headerData: {
          "offerData":{
            name:"项目名称",
            spec:"规格",
            unit:"单位",
            count:"数量",
            unitPrice:"单价",
            total:"金额",
            cardNum:"卡数",
            comments:"备注"
          },
          "offerData1":{
            name:"项目名称1",
            spec:"规格1",
            unit:"单位1",
            count:"数量1",
            unitPrice:"单价1",
            total:"金额1",
            cardNum:"卡数1",
            comments:"备注1"
          }
        },
        // 表格数据
        allData:{
          "offerData":[
            {
              name:"柜体1",
              spec:"sasd",
              unit:"sadad",
              count:"12",
              unitPrice:"14522",
              total:"",
              cardNum:5,
              comments:"大师大师多撒好低"
            },
            {
              name:"柜体2",
              spec:"sasd",
              unit:"sadad",
              count:"12",
              unitPrice:"14522",
              total:"",
              cardNum:5,
              comments:"大师大师多撒好低"
            },
            {
              name:"柜体3",
              spec:"sasd",
              unit:"sadad",
              count:"12",
              unitPrice:"14522",
              total:"",
              cardNum:5,
              comments:"大师大师多撒好低"
            },
            {
              name:"柜体4",
              spec:"sasd",
              unit:"sadad",
              count:"12",
              unitPrice:"14522",
              total:"",
              cardNum:5,
              comments:"大师大师多撒好低"
            },
            {
              name:"柜体5",
              spec:"sasd",
              unit:"sadad",
              count:"12",
              unitPrice:"14522",
              total:"",
              cardNum:5,
              comments:"大师大师多撒好低"
            },
            {
              name:"柜体6",
              spec:"sasd",
              unit:"sadad",
              count:"12",
              unitPrice:"14522",
              total:"",
              cardNum:5,
              comments:"大师大师多撒好低"
            }
          ],
          "offerData1":[
            {
              name:"边框1",
              spec:"sasd",
              unit:"sadad",
              count:"12",
              unitPrice:"14522",
              total:"",
              cardNum:5,
              comments:"大师大师多撒好低"
            },
            {
              name:"边框2",
              spec:"sasd",
              unit:"sadad",
              count:"12",
              unitPrice:"14522",
              total:"",
              cardNum:5,
              comments:"大师大师多撒好低"
            },
            {
              name:"边框3",
              spec:"sasd",
              unit:"sadad",
              count:"12",
              unitPrice:"14522",
              total:"",
              cardNum:5,
              comments:"大师大师多撒好低"
            },
            {
              name:"边框4",
              spec:"sasd",
              unit:"sadad",
              count:"12",
              unitPrice:"14522",
              total:"",
              cardNum:5,
              comments:"大师大师多撒好低"
            },
            {
              name:"边框5",
              spec:"sasd",
              unit:"sadad",
              count:"12",
              unitPrice:"14522",
              total:"",
              cardNum:5,
              comments:"大师大师多撒好低"
            },
            {
              name:"边框6",
              spec:"sasd",
              unit:"sadad",
              count:"12",
              unitPrice:"14522",
              total:"",
              cardNum:5,
              comments:"大师大师多撒好低"
            }
          ]
        },
        //页头部分
        pageHead:[
          {
            name:"订单号：",
            isEdit:false
          },
          {
            name:"客户名称：",
            isEdit:false
          },
          {
            name:"销售日期：",
            isEdit:false
          },
          {
            name:"联系人：",
            isEdit:false
          },
          {
            name:"联系地址：",
            isEdit:false
          }
        ],
        // 页头部分内容
        pageHeadContent:[
          {
            name:""
          },
          {
            name:""
          },
          {
            name:""

          },
          {
            name:""

          },
          {
            name:""

          }
        ],
        options: {
          'offerData':{
            value: 'offerData',
            label: '报价明细单'
          }, 
          'offerData1':{
            value: 'offerData1',
            label: '报价明细单1'
          }
        },
        value: ''

      };
    },
    methods: {
     onSelect(){
      //  console.log(this.options[this.value]['label']);
       this.settingPanelTitle=this.options[this.value]['label'];
     },
     //编辑头部
      editHeader(e){
        if(this.isEdit){
          e.target.innerHTML="保存"
        }else{
          e.target.innerHTML="编辑"
        }
        this.isEdit=!this.isEdit;
      },
     //新增列
      addCol(){
        // this.headerData.push({newCol:"新项"});
       this.$set(this.headerData[this.value], 'newCol'+this.col++, "新列")
      },
      //删除列
      delCol(key){
        this.$delete(this.headerData[this.value],key);
        for(let data of this.allData[this.value]){
          this.$delete(data,key);
        } 
      },
     //导出表格
      onexport(evt){
        let wb=XLSX.utils.table_to_book(document.getElementById("out-table"));
      
        let wbout=XLSX.write(wb,{bookType:'xlsx',type:'binary'});
        
        XLSX_SAVE.saveAs(new Blob([this.s2ab(wbout)],{type:'application/octet-stream'}),"sheetjs.xlsx");
      },
      // 转2进制
      s2ab(s){
        if(typeof ArrayBuffer !== 'undefined'){
          let buf =new ArrayBuffer(s.length);
          let view =new Uint8Array(buf);
          for(var i=0;i!=s.length;++i) view[i] =s.charCodeAt(i) & 0xFF;  
          return buf;      
        }else{
          let buf= new Array(s.length);
          for(var i=0;i!=s.length;++i) buf[i] =s.charCodeAt(i) & 0xFF;
          return buf;
        }
         
      },
      // 拖拽列函数
      dragCol(){
        // 获取th集合
        if(this.$refs.tHead){
          let ths=this.$refs.tHead.querySelectorAll("table>thead th");
          // 阻止拖拽的默认行为
          this.$refs.tHead.ondragover=function(e){
            // console.log("dragover");
            e.preventDefault();
          }
          this.$refs.tHead.ondragstart=function(e){
            // console.log("start");

            //储存所拿起的th元素下标
            e.dataTransfer.setData("obj_add",e.target.cellIndex);
          }
          // 放置元素事件
          this.$refs.tHead.ondrop=function(e){
            e.preventDefault;    
            var i = parseInt(e.dataTransfer.getData("obj_add"));//所拿起的th列下标  
            // console.log(i);  
            var d = e.target.cellIndex;//被放入的列下标

            var _t = e.target; //目前拿起的元素  
            // console.log(_t);
            for(var th of ths){

                //要被交换位置的元素th  
                if(th.cellIndex == i){   
                  //拿起元素和要被放置的元素下标一致，交换位置,判断往前还是往后
                  if(th.cellIndex>_t.cellIndex){
                    th.parentNode.insertBefore(th,_t);
                  }else{
                    th.parentNode.insertBefore(th,_t.nextElementSibling);
                  }
                }  
            } 
            let trs =document.querySelectorAll("table>tbody>tr");
            // console.log(trs);
            for(let tr of trs){

              let drag=""; //拿起的td
              let drop=""; //放下的td
              let tds=tr.children;

              // console.log(tds);
              for(let td of tds){
                
                if(td.cellIndex == i){ 
                    // console.log(td); 
                    drag = td;  
                }  
                if(td.cellIndex == d){ 
                    // console.log(td);  
                    drop = td;  
                } 
                
              }
              if(drag != undefined && drop != undefined && drag != "" && drop != ""){  
                // console.log(drag.parentNode);
                if(drag.cellIndex>drop.cellIndex){
                  drag.parentNode.insertBefore(drag,drop); 
                }else{
                  drag.parentNode.insertBefore(drag,drop.nextElementSibling); 
                }
                
              } 
            }
          }
        }
      },
      //双击编辑事件
      editItem(i){
        this.pageHead[i].isEdit= !this.pageHead[i].isEdit;
      },
      //增加页头项
      addHeadItem(){
        console.log(123);
        this.pageHead.push({name:"新项",isEdit:false});
      },
      // 删除页头项
      delItem(i){
        this.pageHead.splice(i,1);
      },
      // 计算总额
      caclTotal(){
        for(let  key in this.allData){
          for(let data of this.allData[key]){
            data.total=data.unitPrice*data.count;
            this.totalPrice+=data.total;
          }
        }
      },
      //拖动改变表格宽度
      changeColWidth(){
        let tTD={}; //用来存储当前更改宽度的Table Cell   
        let table = this.$refs.content;
        //因为与拖拽的事件冲突，改变宽度目标元素放在tbody行上
        let tr=table.rows[1];  
        for (let j = 0; j < tr.cells.length; j++) {  
       
          tr.cells[j].onmousedown = function () {   
            //记录单元格  
            tTD = this;   
            if (event.offsetX > tTD.offsetWidth - 10) {   
              tTD.mouseDown = true;   
              tTD.oldX = event.x;   
              tTD.oldWidth = tTD.offsetWidth;   
            }    
          };   
          tr.cells[j].onmouseup = function () {   
            //结束宽度调整   
            if (tTD == undefined) tTD = this;   
            tTD.mouseDown = false;   
            tTD.style.cursor = 'default';   
          };   
          tr.cells[j].onmousemove = function () {   
            //更改鼠标样式   
            if (event.offsetX > this.offsetWidth - 10)   
              this.style.cursor = 'col-resize';   
            else   
              this.style.cursor = 'default';   
            //取出暂存的Table Cell   
            if (tTD == undefined) tTD = this;   
            //调整宽度   
            if (tTD.mouseDown) {   
              tTD.style.cursor = 'default';   
              if (tTD.oldWidth + (event.x - tTD.oldX)>0)  
                tTD.width = tTD.oldWidth + (event.x - tTD.oldX);   
              //调整列宽   
              tTD.style.width = tTD.width;   
              tTD.style.cursor = 'col-resize';   
              //调整该列中的每个Cell 
              for (j = 0; j < table.rows.length-1; j++) {  
                table.rows[j].cells[tTD.cellIndex].width = tTD.width;   
              }   
             
            }   
          };   
        }   
      }
    },
    mounted(){
      this.caclTotal();
      
    },
    updated(){
      //改变表头位置
      this.dragCol();
      // console.log(this.pageHeadContent);
      //改变表头宽度
      this.changeColWidth();
    }
  }
</script>

<style scoped>
 

/* 表格样式 start */
  table{
    table-layout: fixed;
    margin-top: 3rem;
    width:100%;
  }
  table th{
    height:4rem;
  }
  td{
    margin-top:1rem ;
    height: 3.5rem;
  }
  #out-table input{
    background:no-repeat 0 0 scroll ＃EEEEEE;
    border:none;
    outline:medium;
    width:40%;
  }
  #out-table{
    margin-top:5rem; 
  }
/* 表格样式 end */
/* 表头定制项样式 start*/
  #out-table li{
    width: 30%;
    margin-top:2rem; 
    position: relative;
  }
  #out-table .pageItem li .close {
    float: none;
    position: absolute;
    top:-14px;
    left:-8px;
  }
/* 表头定制项样式 end*/

/* 设置面板样式 start */
  .setting label{
    font-size: 1.45rem;
    font-weight: normal;
  }


</style>