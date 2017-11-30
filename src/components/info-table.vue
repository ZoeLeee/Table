<template>
  <el-row :style="{ fontSize: settingData[value].fontSize + 'px' }"> 
    <el-col :span="18">
     
      <div id="out-table" v-if="headerData[value]">
        <h1>
          {{settingPanelTitle}}
        </h1>
        <ul class="list-unstyled list-inline pageItem pageHeader" >
          <li v-for="(item,index) in pageHead[value]" :key="item.id">
            <button type="button" class="close" aria-label="Close" v-if="item.isEdit" @click="delItem(index)">
              <span aria-hidden="true">&times;</span>
            </button> 
            <span @dblclick="editItem(index)" v-if="!item.isEdit">{{item.name}}</span>
            <input type="text" class="edit" v-model="item.name" v-if="item.isEdit" @keyup.enter="editItem(index)"> 
            <input type="text" v-model="pageHeadContent[value][index].name">    
          </li>
          <li>
            <el-button type="primary" icon="el-icon-plus" @click="addHeadItem" size="mini"></el-button>
          </li>
        </ul>
        <el-collapse v-model="activeTables">
          <el-collapse-item :title="item" :name="index" v-for="(item,key,index) in showTable" :key="key" draggable="true">
            <div>
              <el-button @click="editHeader($event)" @keyup.enter="editHeader" type="primary">编辑</el-button>
              <el-button type="danger" @click="delTable(key)">删除</el-button>
             
            </div>
            <table class="table" ref="content" :class="{'table-bordered':settingData[value].isBorder,'table-striped':settingData[value].isstriped}" id="mainTable">
              <thead ref="tHead">
                <tr>
                  <template v-for="(item,key1 ) in headerData[value][key]">        
                    <th :key="item.id" draggable="true">
                      <span v-if="isEdit">{{item}}</span> 
                      <input type="text" class="edit" v-model="headerData[value][key][key1]" v-if="!isEdit">
                      <button type="button" class="close" aria-label="Close" @click="delCol(key)" v-if="!isEdit">
                        <span aria-hidden="true">&times;</span>
                      </button> 
                    </th>
                  </template>
                </tr>
              </thead>
              <tbody ref="tbody">
                <tr v-for="data in allData[value][key].data" :key="data.id">
                  <!-- <td v-for="item in headerData[value]" :key="item.id">{{data[getHeadKey(item)]}}</td> -->
                  <td v-for="item in data" :key="item.id">{{item}}</td>
                </tr>
              </tbody>
              <tfoot v-if="allData[value][key].data">
                <tr>
                  <td>
                    合计
                  </td>
                  <td :colspan="Object.keys(headerData[value][key]).length-1">
                    <span>¥{{settingData[value].totalPrice}}</span>
                  </td>
                </tr> 
              </tfoot>
            </table>
          </el-collapse-item>
        </el-collapse> 
         
        <ul class="list-unstyled pull-right">
          <li v-for="item in pageFootContent" :key="item.id">
            {{item.title}}<span>{{item.value}}</span>
          </li>
        </ul>
      </div>
    </el-col>
    <el-col :span="6">
      <el-card class="box-card" body-style="{ padding: 30px }">
        <div slot="header" class="clearfix">
          <el-select v-model="value" placeholder="请选择" @change="onTableSelect" ref="selectedItem">
            <el-option ref="selectedItem"
              v-for="item in options"
              :key="item.index"
              :label="item.label"
              :value="item.value">
            </el-option>
          </el-select>
        </div>
        <el-card>
          <div slot="header" class="clearfix">
            {{settingPanelTitle}}
             <el-button class="pull-right" @click="downloadFile(outputData)" id="export-table" type="success" size="mini">导出</el-button>
             <a id="downlink"></a>
          </div>
          <div id="setting" ref="setting">
            <el-collapse v-model="activeNames" >
              <el-collapse-item name="1" draggable="true">
                <template slot="title">
                  <span>格式设置</span>
                </template>
                <div>
                  <label for="">字体大小:</label>
                  <el-input-number v-model="settingData[value].fontSize" :min="0" size="small" label="字体大小"></el-input-number>
                </div>
                <div>
                  <label for="">边框:</label>
                  <el-switch
                    v-model="settingData[value].isBorder"
                    active-color="#13ce66"
                    inactive-color="#ff4949"
                    :active-value="true"
                    :inactive-value="false">
                  </el-switch>
                </div>
                <div>
                  <label for="">条纹:</label>
                  <el-switch
                    v-model="settingData[value].isstriped"
                    active-color="#13ce66"
                    inactive-color="#ff4949"
                    :active-value="true"
                    :inactive-value="false">
                  </el-switch>
                </div>
              </el-collapse-item>
              <el-collapse-item title="显示设置" name="2" draggable="true"> 
                <el-select v-model="settingData[value].headValue" placeholder="请选择" @change="onTheadSelect">
                  <el-option
                    v-for="(item,key) in headerData[value].summary"
                    :key="item.id"
                    :label="item"
                    :value="key">
                  </el-option>
                </el-select>
                <div class="show-rules">
                  <label for="">大于</label>
                  <el-input-number size="mini" v-model="gtRefValue" :min="0" label="大于参考值"></el-input-number>
                  <el-color-picker v-model="settingData[value].gtBGColor"></el-color-picker>
                </div>
                <div class="show-rules">
                  <label for="">小于</label>
                  <el-input-number size="mini" v-model="ltRefValue" :min="0" label="小于参考值"></el-input-number>
                  <el-color-picker v-model="settingData[value].ltBGColor"></el-color-picker>
                </div>
                <el-button type="primary" @click="addShowRules">添加规则</el-button>
              </el-collapse-item>
              <el-collapse-item title="添加表格" name="5" draggable="true">
                <el-select v-model="selectTable" placeholder="请选择" style="width:130px">
                  <el-option
                    v-for="(item,key) in tableType"
                    :key="key"
                    :label="item"
                    :value="key">
                  </el-option>
                </el-select>
                <el-button icon="el-icon-plus" @click="addTableType"></el-button>
              </el-collapse-item>
              <el-collapse-item title="添加列" name="3" draggable="true">
                <el-input placeholder="请输入内容" v-model="newColName" >
                  <el-select slot="prepend" v-model="selectTable" placeholder="请选择" style="width:130px">
                    <el-option
                      v-for="(item,key) in showTable"
                      :key="key"
                      :label="item"
                      :value="key">
                    </el-option>
                  </el-select>
                  <el-select  slot="prepend" v-model="settingData[value].headValue" placeholder="请选择" @change="onTheadSelect" style="width:130px">
                    <el-option label="空列" :value="0"></el-option>
                    <el-option
                      v-for="(item,key) in headerData[value][selectTable]"
                      :key="item.id"
                      :label="item"
                      :value="key">
                    </el-option>
                  </el-select>
                  <el-button slot="append" icon="el-icon-plus" @click="addCol(newColName)"></el-button>
                </el-input>
              </el-collapse-item>
              <el-collapse-item title="添加统计" name="4" draggable="true">
    
                <el-select v-model="settingData[value].headValue" placeholder="请选择" @change="onCaclSelect" style="width:130px">
    
                  <el-option
                    v-for="(item,key) in headerData[value].summary"
                    :key="item.id"
                    :label="item"
                    :value="key">
                  </el-option>
                </el-select>
                <el-select placeholder="请选择" v-model="caclType" @change="onCaclSelect" style="width:130px">
                  <el-option label="合计" :value="0"></el-option>
                  <el-option label="平均" :value="1"></el-option>
                </el-select>
                <el-button icon="el-icon-plus" @click="addCaclCol"></el-button>
              </el-collapse-item>
              
            </el-collapse>
          
          </div>
        </el-card>
        
      </el-card>
    </el-col> 
  </el-row>
  
</template>

<script>

  import XLSX from 'xlsx';

  export default {
    data() {
      return {
        selectedIndex:"",//选择的列下标
        gtRefValue:0, //参考值 大于时
        ltRefValue:0,
        newColName:"",
        caclType:"",//统计类型
        settingData:{
          'offerData':{
            ltBGColor:"#13ce66",
            gtBGColor:"#ff4949",
            fontSize:14, //字体大小
            headValue:"",
            isBorder:true,
            isstriped:false,
            totalPrice:0//总额

          },
          'offerData1':{
            ltBGColor:"#13ce66",
            gtBGColor:"#ff4949",
            fontSize:14, //字体大小
            headValue:"",
            isBorder:true,
            isstriped:false,
            totalPrice:0
          }
        }, 
        //切换手风琴列表
        activeNames: ['1'],
        activeTables:['0'],
        // 设置面板标题
        settingPanelTitle:"",
        col:1,
        // 是否编辑
        isEdit:true,
        // 表头数据
        headerData: {
          "offerData":{
            summary:{
              name:"项目名称",
              spec:"规格",
              unit:"单位",
              count:"数量",
              unitPrice:"单价",
              total:"金额",
              cardNum:"卡数",
              comments:"备注"
            },
            offerDetail:{
              name:"项目名称",
              spec:"规格",
              unit:"单位",
              count:"数量",
              unitPrice:"单价",
              total:"金额",
              cardNum:"卡数",
      
            },
            proDetail:{
              name:"项目名称",
              spec:"规格",
              unit:"单位",
              count:"数量",
              unitPrice:"单价",
              total:"金额",
           
            }  
          },
          "offerData1":{
            summary:{
              name:"项目名称",
              spec:"规格",
              unit:"单位",
              count:"数量",
              unitPrice:"单价",
              total:"金额",
              cardNum:"卡数",
              comments:"备注"
            },
            offerData1:{
              name:"项目名称",
              spec:"规格",
              unit:"单位",
              count:"数量",
              unitPrice:"单价",
              total:"金额",
              cardNum:"卡数",
              comments:"备注"
            },
            proDetail:{
              name:"项目名称",
              spec:"规格",
              unit:"单位",
              count:"数量",
              unitPrice:"单价",
              total:"金额",
              cardNum:"卡数",
              comments:"备注"
            } 
          }
        },
        // 表格数据
        allData:{
          "offerData":{
            summary:{ 
              title:"报价汇总",
              data:[
                {
                  name:"柜体1",
                  spec:"sasd",
                  unit:"sadad",
                  count:"5",
                  unitPrice:"145",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体2",
                  spec:"sasd",
                  unit:"sadad",
                  count:"10",
                  unitPrice:"522",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体3",
                  spec:"sasd",
                  unit:"sadad",
                  count:"15",
                  unitPrice:"142",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体4",
                  spec:"sasd",
                  unit:"sadad",
                  count:"20",
                  unitPrice:"22",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体5",
                  spec:"sasd",
                  unit:"sadad",
                  count:"25",
                  unitPrice:"222",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体6",
                  spec:"sasd",
                  unit:"sadad",
                  count:"12",
                  unitPrice:"322",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                }
              ]
            },
            offerDetail:{ 
              title:"报价明细",
              data:[
                {
                  name:"柜体1",
                  spec:"sasd",
                  unit:"sadad",
                  count:"5",
                  unitPrice:"145",
                  total:"",
                  cardNum:5
                
                },
                {
                  name:"柜体2",
                  spec:"sasd",
                  unit:"sadad",
                  count:"10",
                  unitPrice:"522",
                  total:"",
                  cardNum:5
                 
                },
                {
                  name:"柜体3",
                  spec:"sasd",
                  unit:"sadad",
                  count:"15",
                  unitPrice:"142",
                  total:"",
                  cardNum:5
                 
                },
                {
                  name:"柜体4",
                  spec:"sasd",
                  unit:"sadad",
                  count:"20",
                  unitPrice:"22",
                  total:"",
                  cardNum:5
                
                },
                {
                  name:"柜体5",
                  spec:"sasd",
                  unit:"sadad",
                  count:"25",
                  unitPrice:"222",
                  total:"",
                  cardNum:5
               
                },
                {
                  name:"柜体6",
                  spec:"sasd",
                  unit:"sadad",
                  count:"12",
                  unitPrice:"322",
                  total:"",
                  cardNum:5
            
                }
              ]
            },
            proDetail:{ 
              title:"产品明细",
              data:[
                {
                  name:"柜体1",
                  spec:"sasd",
                  unit:"sadad",
                  count:"5",
                  unitPrice:"145",
                  total:""
                },
                {
                  name:"柜体2",
                  spec:"sasd",
                  unit:"sadad",
                  count:"10",
                  unitPrice:"522",
                  total:""
                },
                {
                  name:"柜体3",
                  spec:"sasd",
                  unit:"sadad",
                  count:"15",
                  unitPrice:"142",
                  total:""
                },
                {
                  name:"柜体4",
                  spec:"sasd",
                  unit:"sadad",
                  count:"20",
                  unitPrice:"22",
                  total:""
                },
                {
                  name:"柜体5",
                  spec:"sasd",
                  unit:"sadad",
                  count:"25",
                  unitPrice:"222",
                  total:""
                },
                {
                  name:"柜体6",
                  spec:"sasd",
                  unit:"sadad",
                  count:"12",
                  unitPrice:"322",
                  total:""
                }
              ]
            }
          },
          "offerData1":{
            summary:{ 
              title:"报价汇总",
              data:[
                {
                  name:"柜体1",
                  spec:"sasd",
                  unit:"sadad",
                  count:"5",
                  unitPrice:"145",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体2",
                  spec:"sasd",
                  unit:"sadad",
                  count:"10",
                  unitPrice:"522",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体3",
                  spec:"sasd",
                  unit:"sadad",
                  count:"15",
                  unitPrice:"142",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体4",
                  spec:"sasd",
                  unit:"sadad",
                  count:"20",
                  unitPrice:"22",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体5",
                  spec:"sasd",
                  unit:"sadad",
                  count:"25",
                  unitPrice:"222",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体6",
                  spec:"sasd",
                  unit:"sadad",
                  count:"12",
                  unitPrice:"322",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                }
              ]
            },
            offerDetail:{ 
              title:"报价明细",
              data:[
                {
                  name:"柜体1",
                  spec:"sasd",
                  unit:"sadad",
                  count:"5",
                  unitPrice:"145",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体2",
                  spec:"sasd",
                  unit:"sadad",
                  count:"10",
                  unitPrice:"522",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体3",
                  spec:"sasd",
                  unit:"sadad",
                  count:"15",
                  unitPrice:"142",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体4",
                  spec:"sasd",
                  unit:"sadad",
                  count:"20",
                  unitPrice:"22",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体5",
                  spec:"sasd",
                  unit:"sadad",
                  count:"25",
                  unitPrice:"222",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体6",
                  spec:"sasd",
                  unit:"sadad",
                  count:"12",
                  unitPrice:"322",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                }
              ]
            },
            proDetail:{ 
              title:"产品明细",
              data:[
                {
                  name:"柜体1",
                  spec:"sasd",
                  unit:"sadad",
                  count:"5",
                  unitPrice:"145",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体2",
                  spec:"sasd",
                  unit:"sadad",
                  count:"10",
                  unitPrice:"522",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体3",
                  spec:"sasd",
                  unit:"sadad",
                  count:"15",
                  unitPrice:"142",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体4",
                  spec:"sasd",
                  unit:"sadad",
                  count:"20",
                  unitPrice:"22",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体5",
                  spec:"sasd",
                  unit:"sadad",
                  count:"25",
                  unitPrice:"222",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                },
                {
                  name:"柜体6",
                  spec:"sasd",
                  unit:"sadad",
                  count:"12",
                  unitPrice:"322",
                  total:"",
                  cardNum:5,
                  comments:"大师大师多撒好低"
                }
              ]
            }
          }
        },
        //页头部分
        pageHead:{
          "offerData":[
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
          "offerData1":[
            {
              name:"订单号1：",
              isEdit:false
            },
            {
              name:"客户名称1：",
              isEdit:false
            },
            {
              name:"销售日期1：",
              isEdit:false
            },
            {
              name:"联系人1：",
              isEdit:false
            },
            {
              name:"联系地址1：",
              isEdit:false
            }
          ]
        },
        // 页头部分内容
        pageHeadContent:{
          "offerData":[
            {
              name:"1231312313"
            },
            {
              name:"zoe"
            },
            {
              name:"2017-10-21"

            },
            {
              name:"joe"

            },
            {
              name:"beijibng"
            }
          ],
          "offerData1":[
            {
              name:"2231312313"
            },
            {
              name:"zoe1"
            },
            {
              name:"2017-10-21"

            },
            {
              name:"joe11"

            },
            {
              name:"beijibng111"
            }
          ]
        },
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
        //页尾内容
        pageFootContent:[],
        currentPageHead:"",
        value: 'offerData', //选择的表格
        outputData:[], // 导出的数据
        outFile: '',  // 导出文件el
        cloneHead:{},//克隆的表头
        tableType:{}, //表种类
        showTable:{summary:"报价汇总"},
        selectAllData:null, //当前选中的表 
        selectTable:""
      };
    },
    methods: {
      // 选择表格
      onTableSelect(){ 
        this.settingPanelTitle=this.options[this.value]['label'];
      },
      //选择表头项
      onTheadSelect(){
      
        //被选中的表项
        this.selectedIndex=this.settingData[this.value].headValue;
      },
      onCaclSelect(){
        this.selectedIndex=this.settingData[this.value].headValue;
      },
      //获得对应列数据
      getColData(index){
        let selectedData=[];
        let allData=this.allData[this.value].summary.data;
        for(let data of allData){
          selectedData.push(data[index]);
        }
 
        return selectedData;
        
      },
      // 筛选高亮方法
      showRules(index,refValueBig,gtBGColor,refValueSm,ltBGColor){

        //所选中的表项在第几列
        let colIndex=Object.keys(this.allData[this.value].summary.data[0]).indexOf(index);
        let selectedData=this.getColData(index);
        // console.log(selectedData)
        // console.log(this.$refs.tbody);
        for(let i in selectedData){
          
          if(parseFloat(selectedData[i])>refValueBig && refValueBig !==0){
            
            this.$refs.tbody[0].rows[i].cells[colIndex].style.backgroundColor=gtBGColor;
          } 
          
          if(parseFloat(selectedData[i])<refValueSm && refValueSm !==0){
            
            this.$refs.tbody[0].rows[i].cells[colIndex].style.backgroundColor=ltBGColor;
          }
        }
        
      },
      addShowRules(){
       
        this.showRules(this.selectedIndex,this.gtRefValue,this.settingData[this.value].gtBGColor,this.ltRefValue,this.settingData[this.value].ltBGColor); 
      },
      //添加自定义统计
      addCaclCol(){
        
        let selectedData=this.getColData(this.selectedIndex);
        let sum=0;//存储和
        let avg=0;//存储平均值
        if(this.caclType== 0){
          sum = selectedData.reduce(function (a, b) {
            return parseFloat(a) + parseFloat(b);
          });
          this.pageFootContent.push({
            title:this.headerData[this.value].summary[this.selectedIndex]+"合计：",
            value:sum
          });
       
        }else if(this.caclType==1){
          avg=selectedData.reduce(function (a, b) {
            return parseFloat(a) + parseFloat(b);
          })/(selectedData.length);
          this.pageFootContent.push({
            title:this.headerData[this.value].summary[this.selectedIndex]+"平均值：",
            value:avg.toFixed(2)
          })
        }


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
      addCol(colName){
        
          let key='newCol'+ this.col++;
          let allData=this.allData[this.value][this.selectTable].data;
          let headata=this.headerData[this.value][this.selectTable];
          // console.log(key);
          if(!colName){
            if(this.selectedIndex){
              
              colName=headata[this.selectedIndex];
              
            }else{
              colName="新列";
            }
          }
          
          this.$set(headata,key , colName)
          
          if(this.selectedIndex){
            
            //所选中的表项在第几列
            let colIndex=Object.keys(allData[0]).indexOf(this.selectedIndex);
            //获取列数据
            let selectedData=this.getColData(this.selectedIndex);
            
            for(let index in allData){
              
              this.$set(allData[index],key,selectedData[index])
            }
            
          }else{
            for(let index in allData){
              this.$set(allData[index],key,"")
            }
          }
      },
      //删除列
      delCol(key){
        let allData=this.allData[this.value][this.selectTable].data;
        this.$delete(this.headerData[this.value][this.selectTable],key);
        for(let data of allData){
          this.$delete(data,key);
        } 
      },
      
      // 拖拽列函数
      dragCol(){
        // 获取th集合
        
        if(this.$refs.tHead){
          for(let el of this.$refs.tHead){
         
            let ths=el.querySelectorAll("th");
            // 阻止拖拽的默认行为
            el.ondragover=function(e){
              // console.log(el,"dragover");
              e.preventDefault();
            }
            el.ondragstart=function(e){
              // console.log("start");
              //储存所拿起的th元素下标
              e.dataTransfer.setData("obj_add",e.target.cellIndex);
            }
            // 放置元素事件
            el.addEventListener("drop",function(e){
              e.preventDefault;    
              var i = parseInt(e.dataTransfer.getData("obj_add"));//所拿起的th列下标  
              // console.log(i);  
              var d = e.target.cellIndex;//被放入的列下标

              var _t = e.target; //目前拿起的元素  
              // console.log(_t);
              for(var th of ths){
                  //要被交换位置的元素th  
                  if(th.cellIndex == i && e.target.parentNode.nodeName=="TR"){   
                    //拿起元素和要被放置的元素下标一致，交换位置,判断往前还是往后

                    if(th.cellIndex>_t.cellIndex){
                      // console.log(th);
                      // console.log(e.target.parentNode.nodeName);
                      e.target.parentNode.insertBefore(th,_t);
                    }else{
                      e.target.parentNode.insertBefore(th,_t.nextElementSibling);
                    }
                  }  
              } 
              let trs =el.parentNode.querySelectorAll("tbody>tr");
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
            },false)
          }
        }
      },
      //双击编辑事件
      editItem(i){
        // console.log(i);
        this.pageHead[this.value][i].isEdit= !this.pageHead[this.value][i].isEdit;
      },
      //增加页头项
      addHeadItem(){
        
        this.pageHead[this.value].push({name:"新项:",isEdit:false});
        this.pageHeadContent[this.value].push({name:""});
        
      },
      // 删除页头项
      delItem(i){
        this.pageHead[this.value].splice(i,1);
      },
      // 计算总额
      caclTotal(){
        for(let key in this.allData){
     
          for(let data of this.allData[key].summary.data){
            data.total=data.unitPrice*data.count;
            this.settingData[key].totalPrice+=parseFloat(data.total);
          }
        }
      },
      //拖动改变表格宽度
      changeColWidth(){
        let tTD={}; //用来存储当前更改宽度的Table Cell   
        let tables = this.$refs.content;
        for(let table of tables){
          if(!table){
            table=document.getElementById("mainTable");
          }
          // console.log(table);
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
      // 导出功能
      downloadFile(rs) { // 点击导出按钮
        //拼接导出的数据
        //1.拼接标题
        rs=[];
        rs.push({title:this.settingPanelTitle});

        //2.拼接表头内容
        
        let headContent={};
        //页头项
        let pHead=this.pageHead[this.value];
        //页头项对应的内容
        let pHeadContent=this.pageHeadContent[this.value];

        for(let i in pHead){  
          if(i %3==0){
            rs.push(headContent);
            headContent={};
          } 
          headContent[i]=pHead[i].name+""+pHeadContent[i].name;
          
        }
        // console.log(headContent);
        rs.push(headContent);
        rs.push({});//excle表格换行
        
        for(let index in this.showTable){
        //3.拼接表头标题
          let title={};
          title[index]=this.showTable[index];
          rs.push({}); //excle表格换行
          rs.push(title); 
          rs.push(this.headerData[this.value][index]);

          //4.拼接表内容
          rs.push(...this.allData[this.value][index].data);
        }
        //5.拼接表尾
        rs.push({title:"合计",value:this.settingData[this.value].totalPrice});
        // 6.拼接页尾
        let pageFoot={};
        
        for(let index in this.pageFootContent){
          pageFoot[index]=this.pageFootContent[index].title+this.pageFootContent[index].value;
        }
        rs.push(pageFoot);
        
        this.downloadExl(rs, this.settingPanelTitle)
      },
      downloadExl(json, downName, type) {  // 导出到excel
        
        let tmpdata = [] // 用来保存转换好的json
        const wopts = { bookType: 'xlsx', bookSST: true, type: 'binary', cellStyles: true };
        let maxLen=0; //最长的一行

        var ws = { };  
        ws['!cols']= [];
        json.map((v, i) => {
          if(maxLen<Object.keys(v).length){
            maxLen=Object.keys(v).length;
          }
          return Object.keys(v).map((k, j) => {  //取出键对应的值

            return Object.assign({}, { //拼接输出的sheet
              v: v[k],
              position: (j > 25 ? this.getCharCol(j) : String.fromCharCode(65 + j)) + (i + 1),
               
            })}
          )
        }).reduce((prev, next) =>  prev.concat(next)).forEach(function (v) {
          var cell={ //转换输出json
            v: v.v
          }
          if ( typeof cell.v === 'number')  
               cell.t = 'n';  
            else if ( typeof cell.v === 'boolean')  
                cell.t = 'b';  
            else if (cell.v instanceof Date) {  
                cell.t = 'n';  
                cell.z = XLSX.SSF._table[14];  
                cell.v = datenum(cell.v);  
            } else  
                cell.t = 's';
          tmpdata[v.position] = cell;
          
        })
        // console.log(tmpdata);
        let outputPos = Object.keys(tmpdata) ; // 设置区域,比如表格从A1到D10
        // console.log( tmpdata["A8"]);
        tmpdata["A8"].s = { font: { sz: 14, bold: true, color: { rgb: "FFFFAA00" } }, fill: { bgColor: { indexed: 64 }, fgColor: { rgb: "FFFF00" } } };//<====设置xlsx单元格样式
        tmpdata["!merges"] = [{
                s: { c: 0, r: 0 },
                e: { c: 4, r: 0 }
            }];//<====合并单元格 
        for(var n = 0; n != maxLen; ++n){  
          ws['!cols'].push({  
            wpx: 170  
          });  
        }
        // 转化最长的行所对应的区域码
        maxLen=this.getCharCol(maxLen);
        
        // console.log(maxLen+outputPos[outputPos.length-1].slice(1));
        let tmpWB = {
          SheetNames: ['mySheet'], // 保存的表标题
          Sheets: {
            'mySheet': Object.assign({},
              tmpdata, // 内容
              ws,
              {cellStyles: true},
              {
                '!ref': outputPos[0] + ':' + (maxLen+outputPos[outputPos.length-1].slice(1)) // 设置填充区域
              })
          }
        }
        
        let tmpDown = new Blob([this.s2ab(XLSX.write(tmpWB,
          {cellStyles: true,bookType: (type === undefined ? 'xlsx' : type), bookSST: false, type: 'binary'} // 这里的数据是用来定义导出的格式类型
        ))], {
          type: ''
        })  // 创建二进制对象写入转换好的字节流
        var href = URL.createObjectURL(tmpDown)  // 创建对象超链接
        this.outFile.download = downName + '.xlsx'  // 下载名称
        this.outFile.href = href  // 绑定a标签
        this.outFile.click()  // 模拟点击实现下载
        setTimeout(function () {  // 延时释放
          URL.revokeObjectURL(tmpDown) // 用URL.revokeObjectURL()来释放这个object URL
        }, 100)
     
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
      // 将指定的自然数转换为26进制表示。映射关系：[0-25] -> [A-Z]。
      getCharCol(n) { 
        let s = ''
        let m = 0
        while (n > 0) {
          m = n % 26 + 1
          s = String.fromCharCode(m + 64) + s
          n = (n - m) / 26
        }
        return s
      },
      //由表头的value获取key
      getHeadKey(key){
 
        // 交换key和value位置后的新对象
        let newHead={};
        let headerData=this.cloneHead[this.value]
        for(let i in headerData){
          // console.log(i,headerData[i]);
          newHead[headerData[i]]=i;
        }
        // console.log(newHead);
        return newHead[key];
      },
      // 获取表格类型数据
      getTableType(){
        this.tableType={};
        let allData=this.allData[this.value];
        
        for (let index in allData){
          let tmpObj={};
          this.tableType[index]=allData[index].title;
        }
      },
      addTableType(){

        if(this.selectTable){
          this.$set(this.showTable,this.selectTable,this.tableType[this.selectTable]);
        }
      },
      delTable(key){
        // console.log(key);
        this.$delete(this.showTable,key);
      },
      //拖拽设置面板
      dragSettingPanel(){
        // console.log()
        //获得设置面板对象
        let Panels=document.querySelectorAll(".el-collapse");
        // console.log(setPanel)
        // 拿起的元素
        let dragElem={};
        // console.log(setPanel)
        Panels.forEach((item)=>{
          //消除拖拽默认行为
          item.ondragover=function(e){
            // console.log("dragover")
            e.preventDefault();
          }
          item.ondragstart=function(e){
            // console.log("start");
            //储存所拿起的元素
            dragElem=e.target;
          }
          item.ondrop=function(e){
            //储存放入的目标元素
            let dropElem=e.target;

            while(dropElem.className != "el-collapse-item" && dropElem.className != "el-collapse-item is-active"){
              dropElem=dropElem.parentNode;
            }
            // console.log(dropElem); 
            this.insertBefore(dragElem,dropElem);
          }
        })
        
      }
    },
    mounted(){
      // this.getHeadKey();
      this.cloneHead=JSON.parse(JSON.stringify(this.headerData));;
      this.currentPageHead=this.pageHead[this.value];
      this.caclTotal();
      this.onTableSelect();
      this.outFile = document.getElementById('downlink');
      this.getTableType();
      this.dragSettingPanel();
    },
    updated(){
      //改变表头位置
      this.dragCol();
      //改变表头宽度
      this.changeColWidth();
  
    }
    
  }
</script>

<style scoped>
  /* 编辑框 */
  #out-table .edit{
    border:1px dashed #dddddd
  }
  h1{
    text-align: center
  }

/* 表格样式 start */
  table{
    table-layout: fixed;
    width:100%;
    margin-top:2.5rem; 
  }
  table th{
    height:4rem;
  }
  td{
    margin-top:1rem ;
    height: 3.5rem;
  }
  #out-table input{
    border:none;
    outline:medium;
    width:25%;
  }
  #out-table{
    margin-top:5rem; 
  }
/* 表格样式 end */
/* 表头定制项样式 start*/
  #out-table .pageHeader li{
    width: 30%;
    margin-top:2rem; 
    position: relative;
  }
  #out-table .pageItem li .close {
    float: none;
    position: absolute;
    top:-10px;
    left:75px;
  }
/* 表头定制项样式 end*/

/* 设置面板样式 start */
  #setting label{
    font-size: 1.45rem;
    font-weight: normal;
    margin-right:2rem; 
    vertical-align: sub;
  }
  #setting div{
    vertical-align: middle;
  }
  .show-rules{
    margin:2rem;
  }
  .el-input-group__prepend {
    width: auto;
  }
</style>