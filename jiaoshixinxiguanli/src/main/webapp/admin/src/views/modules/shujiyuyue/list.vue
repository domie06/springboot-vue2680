<template>
    <div class="main-content">
        <el-dialog title="审核" :visible.sync="shujiyuyueYesnoTypesVisible">
            <el-form :model="form">
                <input type="hidden" v-model="form.id">
                <el-form-item label="审核" >
                    <el-select v-model="form.shujiyuyueYesnoTypes" placeholder="请选择审核类型">
                        <el-option label="通过" value="2"></el-option>
                        <el-option label="拒绝" value="3"></el-option>
                    </el-select>
                </el-form-item>
                <el-form-item label="审核意见">
                    <el-input type="textarea" v-model="form.shujiyuyueYesnoText" placeholder="审核意见"></el-input>
                </el-form-item>
            </el-form>
            <div slot="footer" class="dialog-footer">
                <el-button @click="shujiyuyueYesnoTypesVisible = false">取 消</el-button>
                <el-button type="primary" @click="shujiyuyueYesnoTypesShenhe">提 交</el-button>
            </div>
        </el-dialog>
        <!-- 条件查询 -->
        <div v-if="showFlag">
            <el-form :inline="true" :model="searchForm" class="form-content">
                <el-row :gutter="20" class="slt" :style="{justifyContent:contents.searchBoxPosition=='1'?'flex-start':contents.searchBoxPosition=='2'?'center':'flex-end'}">
                                                                                
                                                             
                    <el-form-item :label="contents.inputTitle == 1 ? '教室标题' : ''">
                        <el-input prefix-icon="el-icon-search" v-model="searchForm.jiaoshiName" placeholder="教室标题" clearable></el-input>
                    </el-form-item>
                     
                    <el-form-item :label="contents.inputTitle == 1 ? '教室类型' : ''">
                        <el-select v-model="searchForm.jiaoshiTypes" placeholder="请选择教室类型">
                            <el-option label="=-请选择-=" value=""></el-option>
                            <el-option
                                    v-for="(item,index) in jiaoshiTypesSelectSearch"
                                    v-bind:key="index"
                                    :label="item.indexName"
                                    :value="item.codeIndex">
                                <!--lable是要展示的名称-->
                                <!--value是值-->
                            </el-option>
                        </el-select>
                    </el-form-item>
                     
                    <el-form-item :label="contents.inputTitle == 1 ? '教室状态' : ''">
                        <el-select v-model="searchForm.jiaoshiStatusTypes" placeholder="请选择教室状态">
                            <el-option label="=-请选择-=" value=""></el-option>
                            <el-option
                                    v-for="(item,index) in jiaoshiStatusTypesSelectSearch"
                                    v-bind:key="index"
                                    :label="item.indexName"
                                    :value="item.codeIndex">
                                <!--lable是要展示的名称-->
                                <!--value是值-->
                            </el-option>
                        </el-select>
                    </el-form-item>
                                                                                                                                                                 
                    <el-form-item :label="contents.inputTitle == 1 ? '教师姓名' : ''">
                        <el-input prefix-icon="el-icon-search" v-model="searchForm.laoshiName" placeholder="教师姓名" clearable></el-input>
                    </el-form-item>
                                                                                                                                             
                    <el-form-item :label="contents.inputTitle == 1 ? '用户姓名' : ''">
                        <el-input prefix-icon="el-icon-search" v-model="searchForm.yonghuName" placeholder="用户姓名" clearable></el-input>
                    </el-form-item>
                                                                                                                        

                    <el-form-item>
                        <el-button type="success" @click="search()">查询<i class="el-icon-search el-icon--right"/></el-button>
                    </el-form-item>
                </el-row>
                <el-row class="ad" :style="{justifyContent:contents.btnAdAllBoxPosition=='1'?'flex-start':contents.btnAdAllBoxPosition=='2'?'center':'flex-end'}">
                    <el-form-item>
                        <el-button
                                v-if="isAuth('shujiyuyue','新增')"
                                type="success"
                                icon="el-icon-plus"
                                @click="addOrUpdateHandler()"
                        >预约教室</el-button>
                        &nbsp;
                        <el-button
                                v-if="isAuth('shujiyuyue','删除')"
                                :disabled="dataListSelections.length <= 0"
                                type="danger"
                                icon="el-icon-delete"
                                @click="deleteHandler()"
                        >删除</el-button>
                        &nbsp;
                        <el-button
                                v-if="isAuth('shujiyuyue','报表')"
                                type="success"
                                icon="el-icon-pie-chart"
                                @click="chartDialog()"
                        >报表</el-button>
                        &nbsp;
                        <a style="text-decoration:none" class="el-button el-button--success"
                           v-if="isAuth('shujiyuyue','导入导出')"
                           icon="el-icon-download"
                           href="http://localhost:8080/jiaoshixinxiguanli/upload/shujiyuyueMuBan.xls"
                        >批量导入教室预约数据模板</a>
                        &nbsp;
                        <el-upload
                                v-if="isAuth('shujiyuyue','导入导出')"
                                style="display: inline-block"
                                action="jiaoshixinxiguanli/file/upload"
                                :on-success="shujiyuyueUploadSuccess"
                                :on-error="shujiyuyueUploadError"
                                :show-file-list = false>
                            <el-button
                                    v-if="isAuth('shujiyuyue','导入导出')"
                                    type="success"
                                    icon="el-icon-upload2"
                            >批量导入教室预约数据</el-button>
                        </el-upload>
                        &nbsp;
                        <!-- 导出excel -->
                        <download-excel v-if="isAuth('shujiyuyue','导入导出')" style="display: inline-block" class = "export-excel-wrapper" :data = "dataList" :fields = "json_fields" name = "shujiyuyue.xls">
                            <!-- 导出excel -->
                            <el-button
                                    type="success"
                                    icon="el-icon-download"
                            >导出</el-button>
                        </download-excel>
                        &nbsp;
                    </el-form-item>
                </el-row>
            </el-form>
            <div class="table-content">
                <el-table class="tables" :size="contents.tableSize" :show-header="contents.tableShowHeader"
                          :header-row-style="headerRowStyle" :header-cell-style="headerCellStyle"
                          :border="contents.tableBorder"
                          :fit="contents.tableFit"
                          :stripe="contents.tableStripe"
                          :row-style="rowStyle"
                          :cell-style="cellStyle"
                          :style="{width: '100%',fontSize:contents.tableContentFontSize,color:contents.tableContentFontColor}"
                          v-if="isAuth('shujiyuyue','查看')"
                          :data="dataList"
                          v-loading="dataListLoading"
                          @selection-change="selectionChangeHandler">
                    <el-table-column  v-if="contents.tableSelection"
                                      type="selection"
                                      header-align="center"
                                      align="center"
                                      width="50">
                    </el-table-column>
                    <el-table-column label="索引" v-if="contents.tableIndex" type="index" width="50" />
					<el-table-column  :sortable="contents.tableSortable" :align="contents.tableAlign"
					                  header-align="center"
					                  label="身份">
					    <template slot-scope="scope">
					        {{scope.row.laoshiId!=null?"教师":scope.row.yonghuId!=null?"用户":""}}
					    </template>
					</el-table-column>
					<el-table-column  :sortable="contents.tableSortable" :align="contents.tableAlign"
					                  header-align="center"
					                  label="姓名">
					    <template slot-scope="scope">
					        {{scope.row.laoshiId!=null?scope.row.laoshiName:scope.row.yonghuId!=null?scope.row.yonghuName:""}}
					    </template>
					</el-table-column>
                    <el-table-column  :sortable="contents.tableSortable" :align="contents.tableAlign"
                                      prop="jiaoshiUuidNumber"
                                      header-align="center"
                                      label="教室编号">
                        <template slot-scope="scope">
                            {{scope.row.jiaoshiUuidNumber}}
                        </template>
                    </el-table-column>
                    <el-table-column  :sortable="contents.tableSortable" :align="contents.tableAlign"
                                      prop="jiaoshiName"
                                      header-align="center"
                                      label="教室标题">
                        <template slot-scope="scope">
                            {{scope.row.jiaoshiName}}
                        </template>
                    </el-table-column>
                    <el-table-column  :sortable="contents.tableSortable" :align="contents.tableAlign"
                                      prop="jiaoshiTypes"
                                      header-align="center"
                                      label="教室类型">
                        <template slot-scope="scope">
                            {{scope.row.jiaoshiValue}}
                        </template>
                    </el-table-column>
                    <el-table-column  :sortable="contents.tableSortable" :align="contents.tableAlign"
                                      prop="jiaoshiStatusTypes"
                                      header-align="center"
                                      label="教室状态">
                        <template slot-scope="scope">
                            {{scope.row.jiaoshiStatusValue}}
                        </template>
                    </el-table-column>
                    <el-table-column :sortable="contents.tableSortable" :align="contents.tableAlign" prop="jiaoshiPhoto"
                               header-align="center"
                               width="200"
                               label="教室图片">
                        <template slot-scope="scope">
                            <div v-if="scope.row.jiaoshiPhoto">
                                <img :src="scope.row.jiaoshiPhoto" width="100" height="100">
                            </div>
                            <div v-else>无图片</div>
                        </template>
                    </el-table-column>
                    <el-table-column  :sortable="contents.tableSortable" :align="contents.tableAlign"
                                      prop="jiaoshiAddress"
                                      header-align="center"
                                      label="教室地址">
                        <template slot-scope="scope">
                            {{scope.row.jiaoshiAddress}}
                        </template>
                    </el-table-column>
                    <el-table-column  :sortable="contents.tableSortable" :align="contents.tableAlign"
                                      prop="jiaoshiRongnaliang"
                                      header-align="center"
                                      label="教室容纳量">
                        <template slot-scope="scope">
                            {{scope.row.jiaoshiRongnaliang}}
                        </template>
                    </el-table-column>
                    <el-table-column  :sortable="contents.tableSortable" :align="contents.tableAlign"
                                      prop="shujiyuyueText"
                                      header-align="center"
                                      label="预约信息">
                        <template slot-scope="scope">
                            <span v-if="scope.row.shujiyuyueText != null &&scope.row.shujiyuyueText.length>10">
                                {{scope.row.shujiyuyueText.slice(0,10)}}...
                            </span>
                            <span v-else>
                                {{scope.row.shujiyuyueText}}
                            </span>
                        </template>
                    </el-table-column>
                    <el-table-column  :sortable="contents.tableSortable" :align="contents.tableAlign"
                                      prop="shujiyuyueYesnoTypes"
                                      header-align="center"
                                      label="预约审核">
                        <template slot-scope="scope">
                            {{scope.row.shujiyuyueYesnoValue}}
                        </template>
                    </el-table-column>
                    <el-table-column  :sortable="contents.tableSortable" :align="contents.tableAlign"
                                      prop="shujiyuyueYesnoText"
                                      header-align="center"
                                      label="审核原因">
                        <template slot-scope="scope">
                            <span v-if="scope.row.shujiyuyueYesnoText != null &&scope.row.shujiyuyueYesnoText.length>10">
                                {{scope.row.shujiyuyueYesnoText.slice(0,10)}}...
                            </span>
                            <span v-else>
                                {{scope.row.shujiyuyueYesnoText}}
                            </span>
                        </template>
                    </el-table-column>

                    <el-table-column  :sortable="contents.tableSortable" :align="contents.tableAlign"
                                   prop="shujiyuyueTime"
                                   header-align="center"
                                   label="预约时间">
                        <template slot-scope="scope">
                            {{scope.row.shujiyuyueTime}}
                        </template>
                    </el-table-column>

                    <el-table-column  :sortable="contents.tableSortable" :align="contents.tableAlign"
                                   prop="insertTime"
                                   header-align="center"
                                   label="添加时间">
                        <template slot-scope="scope">
                            {{scope.row.insertTime}}
                        </template>
                    </el-table-column>

                    <el-table-column width="300" :align="contents.tableAlign"
                                     header-align="center"
                                     label="操作">
                        <template slot-scope="scope">
                            <el-button v-if="isAuth('shujiyuyue','查看')" type="success" icon="el-icon-tickets" size="mini" @click="addOrUpdateHandler(scope.row.id,'info')">详情</el-button>
                            <el-button v-if="isAuth('shujiyuyue','修改')" type="primary" icon="el-icon-edit" size="mini" @click="addOrUpdateHandler(scope.row.id)">修改</el-button>
                            <el-button v-if="isAuth('shujiyuyue','删除')" type="danger" icon="el-icon-delete" size="mini" @click="deleteHandler(scope.row.id)">删除</el-button>
                            <el-button v-if="isAuth('shujiyuyue','审核') && scope.row.shujiyuyueYesnoTypes == 1 " type="primary" icon="el-icon-thumb" size="mini" @click="openYesnoTypes(scope.row.id,scope.row.jiaoshiId)">审核</el-button>

                        </template>
                    </el-table-column>
                </el-table>
                <el-pagination
                        clsss="pages"
                        :layout="layouts"
                        @size-change="sizeChangeHandle"
                        @current-change="currentChangeHandle"
                        :current-page="pageIndex"
                        :page-sizes="[10, 20, 50, 100]"
                        :page-size="Number(contents.pageEachNum)"
                        :total="totalPage"
                        :small="contents.pageStyle"
                        class="pagination-content"
                        :background="contents.pageBtnBG"
                        :style="{textAlign:contents.pagePosition==1?'left':contents.pagePosition==2?'center':'right'}"
                ></el-pagination>
            </div>
        </div>
        <!-- 添加/修改页面  将父组件的search方法传递给子组件-->
        <add-or-update v-if="addOrUpdateFlag" :parent="this" ref="addOrUpdate"></add-or-update>



        <el-dialog title="统计报表" :visible.sync="chartVisiable" width="800">
            <el-date-picker
                    v-model="echartsDate"
                    type="year"
                    placeholder="选择年">
            </el-date-picker>
            <el-button @click="chartDialog()">查询</el-button>
            <div id="statistic" style="width:100%;height:600px;"></div>
            <span slot="footer" class="dialog-footer">
				<el-button @click="chartVisiable = false">返回</el-button>
			</span>
        </el-dialog>

    </div>
</template>
<script>
    import AddOrUpdate from "./add-or-update";
    import styleJs from "../../../utils/style.js";

    export default {
        data() {
        return {
            searchForm: {
                key: ""
            },
            sessionTable : "",//登录账户所在表名
            role : "",//权限
            userId:"",//当前登录人的id
    //级联表下拉框搜索条件
              jiaoshiTypesSelectSearch : [],
              jiaoshiStatusTypesSelectSearch : [],
    //当前表下拉框搜索条件
            form:{
                id : null,
                jiaoshiId : null,
                yonghuId : null,
                laoshiId : null,
                shujiyuyueText : null,
                shujiyuyueYesnoTypes : null,
                shujiyuyueYesnoText : null,
                shujiyuyueTime : null,
                insertTime : null,
                createTime : null,
            },
            dataList: [],
            pageIndex: 1,
            pageSize: 10,
            totalPage: 0,
            dataListLoading: false,
            dataListSelections: [],
            showFlag: true,
            sfshVisiable: false,
            shForm: {},
            chartVisiable: false,
            echartsDate: new Date(),//echarts的时间查询字段
            addOrUpdateFlag:false,
            contents:null,
            layouts: '',

            shujiyuyueYesnoTypesVisible : false,

            //导出excel
            json_fields: {
                //级联表字段
                     '教室编号': 'jiaoshiUuidNumber',
                     '教室标题': 'jiaoshiName',
                     '教室类型': 'jiaoshiTypes',
                     '教室状态': 'jiaoshiStatusTypes',
                     '教室图片': 'jiaoshiPhoto',
                     '教室地址': 'jiaoshiAddress',
                     '教室容纳量': 'jiaoshiRongnaliang',
                     '教师姓名': 'laoshiName',
                     '头像': 'laoshiPhoto',
                     '联系方式': 'laoshiPhone',
                     '邮箱': 'laoshiEmail',
                     '用户姓名': 'yonghuName',
                     '头像': 'yonghuPhoto',
                     '联系方式': 'yonghuPhone',
                     '用户身份证号': 'yonghuIdNumber',
                     '邮箱': 'yonghuEmail',
                //本表字段
                     '预约信息': "shujiyuyueText",
                     '预约审核': "shujiyuyueYesnoTypes",
                     '审核原因': "shujiyuyueYesnoText",
                     '预约时间': "shujiyuyueTime",
                     '添加时间': "insertTime",
            },

            };
        },
        created() {
            this.contents = styleJs.listStyle();
            this.init();
            this.getDataList();
            this.contentStyleChange()
        },
        mounted() {
            //获取当前登录用户的信息
            this.sessionTable = this.$storage.get("sessionTable");
            this.role = this.$storage.get("role");
            this.userId = this.$storage.get("userId");

        },
        filters: {
            htmlfilter: function (val) {
                return val.replace(/<[^>]*>/g).replace(/undefined/g,'');
            }
        },
        components: {
            AddOrUpdate,
        },
        computed: {
        },
        methods: {
            chartDialog() {
                let _this = this;
                let params = {
                    dateFormat :"%Y", //%Y-%m
                    riqi :_this.echartsDate.getFullYear(),
                    thisTable : {//当前表
                        tableName :"shujiyuyue",//当前表表名,
                        sumColum : 'shujiyuyue_number', //求和字段
                        date : 'insert_time',//分组日期字段
                        // string : 'shujiyuyue_name',//分组字符串字段
                        // types : 'shujiyuyue_types',//分组下拉框字段
                    },
                    // joinTable : {//级联表（可以不存在）
                    //     tableName :"yonghu",//级联表表名
                    //     // date : 'insert_time',//分组日期字段
                    //     string : 'yonghu_name',//分组字符串字段
                    //     // types : 'yonghu_types',//分组下拉框字段
                    // }
                }
                _this.chartVisiable = true;
                _this.$nextTick(() => {
                    var statistic = this.$echarts.init(document.getElementById("statistic"), 'macarons');
                    this.$http({
                        url: "barSum",
                        method: "get",
                        params: params
                    }).then(({data}) => {
                        if(data && data.code === 0){

                            //柱状图 求和 已成功使用
                            //start
                            let yAxisName = "数值";//根据查询数据具体改(单列要改,多列不改)
                            let xAxisName = "月份";
                            let series = [];//具体数据值
                            data.data.yAxis.forEach(function (item,index) {
                                let tempMap = {};
                                // tempMap.name= ["数值"];//根据查询数据具体改(单列要改,多列不改)
                                tempMap.name=data.data.legend[index];
                                tempMap.type='bar';
                                tempMap.data=item;
                                series.push(tempMap);

                            })

                            var option = {
                                tooltip: {
                                    trigger: 'axis',
                                    axisPointer: {
                                        type: 'cross',
                                        crossStyle: {
                                            color: '#999'
                                        }
                                    }
                                },
                                toolbox: {
                                    feature: {
                                        // dataView: { show: true, readOnly: false },  // 数据查看
                                        magicType: { show: true, type: ['line', 'bar'] },//切换图形展示方式
                                        // restore: { show: true }, // 刷新
                                        saveAsImage: { show: true }//保存
                                    }
                                },
                                legend: {
                                    data: data.data.legend//标题  可以点击导致某一列数据消失
                                },
                                xAxis: [
                                    {
                                        type: 'category',
                                        name: xAxisName,
                                        data: data.data.xAxis,
                                        axisPointer: {
                                            type: 'shadow'
                                        }
                                    }
                                ],
                                yAxis: [
                                    {
                                        type: 'value',//不能改
                                        name: yAxisName,//y轴单位
                                        axisLabel: {
                                            formatter: '{value}' // 后缀
                                        }
                                    }
                                ],
                                series:series//具体数据
                            };
                            // 使用刚指定的配置项和数据显示图表。
                            statistic.setOption(option,true);
                            //根据窗口的大小变动图表
                            window.onresize = function () {
                                statistic.resize();
                            };
                            //end
                        }else {
                            this.$message({
                                message: "报表未查询到数据",
                                type: "success",
                                duration: 1500,
                                onClose: () => {
                                    this.search();
                                }
                            });
                        }
                    });
                });
                ////饼状图
                //_this.chartVisiable = true;
                // this.$nextTick(()=>{
                //     var statistic = this.$echarts.init(document.getElementById("statistic"),'macarons');
                //     let params = {
                //         tableName: "shujiyuyue",
                //         groupColumn: "shujiyuyue_types",
                //     }
                //     this.$http({
                //         url: "newSelectGroupCount",
                //         method: "get",
                //         params: params
                //     }).then(({data}) => {
                //         if (data && data.code === 0) {
                //             let res = data.data;
                //             let xAxis = [];
                //             let yAxis = [];
                //             let pArray = []
                //             for(let i=0;i<res.length;i++){
                //                 xAxis.push(res[i].name);
                //                 yAxis.push(res[i].value);
                //                 pArray.push({
                //                     value: res[i].value,
                //                     name: res[i].name
                //                 })
                //                 var option = {};
                //                 option = {
                //                     title: {
                //                         text: '保险合同类型统计',
                //                         left: 'center'
                //                     },
                //                     tooltip: {
                //                         trigger: 'item',
                //                         formatter: '{b} : {c} ({d}%)'
                //                     },
                //                     series: [
                //                         {
                //                             type: 'pie',
                //                             radius: '55%',
                //                             center: ['50%', '60%'],
                //                             data: pArray,
                //                             emphasis: {
                //                                 itemStyle: {
                //                                     shadowBlur: 10,
                //                                     shadowOffsetX: 0,
                //                                     shadowColor: 'rgba(0, 0, 0, 0.5)'
                //                                 }
                //                             }
                //                         }
                //                     ]
                //                 };
                //                 statistic.setOption(option);
                //                 window.onresize = function() {
                //                     statistic.resize();
                //                 };
                //             }
                //         }
                //     });
                // })
            },
            contentStyleChange() {
                this.contentSearchStyleChange()
                this.contentBtnAdAllStyleChange()
                this.contentSearchBtnStyleChange()
                this.contentTableBtnStyleChange()
                this.contentPageStyleChange()
            },
            contentSearchStyleChange() {
                this.$nextTick(() => {
                    document.querySelectorAll('.form-content .slt .el-input__inner').forEach(el => {
                        let textAlign = 'left'
                        if(this.contents.inputFontPosition == 2)
                            textAlign = 'center'
                            if (this.contents.inputFontPosition == 3) textAlign = 'right'
                                el.style.textAlign = textAlign
                            el.style.height = this.contents.inputHeight
                            el.style.lineHeight = this.contents.inputHeight
                            el.style.color = this.contents.inputFontColor
                            el.style.fontSize = this.contents.inputFontSize
                            el.style.borderWidth = this.contents.inputBorderWidth
                            el.style.borderStyle = this.contents.inputBorderStyle
                            el.style.borderColor = this.contents.inputBorderColor
                            el.style.borderRadius = this.contents.inputBorderRadius
                            el.style.backgroundColor = this.contents.inputBgColor
                    })
                    if (this.contents.inputTitle) {
                        document.querySelectorAll('.form-content .slt .el-form-item__label').forEach(el => {
                            el.style.color = this.contents.inputTitleColor
                            el.style.fontSize = this.contents.inputTitleSize
                            el.style.lineHeight = this.contents.inputHeight
                        })
                    }
                    setTimeout(() => {
                        document.querySelectorAll('.form-content .slt .el-input__prefix').forEach(el => {
                            el.style.color = this.contents.inputIconColor
                            el.style.lineHeight = this.contents.inputHeight
                        })
                        document.querySelectorAll('.form-content .slt .el-input__suffix').forEach(el => {
                            el.style.color = this.contents.inputIconColor
                            el.style.lineHeight = this.contents.inputHeight
                        })
                        document.querySelectorAll('.form-content .slt .el-input__icon').forEach(el => {
                            el.style.lineHeight = this.contents.inputHeight
                        })
                    }, 10 )
                })
            },
            // 搜索按钮
            contentSearchBtnStyleChange() {
                this.$nextTick(() => {
                    document.querySelectorAll('.form-content .slt .el-button--success').forEach(el => {
                        el.style.height = this.contents.searchBtnHeight
                        el.style.color = this.contents.searchBtnFontColor
                        el.style.fontSize = this.contents.searchBtnFontSize
                        el.style.borderWidth = this.contents.searchBtnBorderWidth
                        el.style.borderStyle = this.contents.searchBtnBorderStyle
                        el.style.borderColor = this.contents.searchBtnBorderColor
                        el.style.borderRadius = this.contents.searchBtnBorderRadius
                        el.style.backgroundColor = this.contents.searchBtnBgColor
                    })
                })
            },
            // 新增、批量删除
            contentBtnAdAllStyleChange() {
                this.$nextTick(() => {
                    document.querySelectorAll('.form-content .ad .el-button--success').forEach(el => {
                        el.style.height = this.contents.btnAdAllHeight
                        el.style.color = this.contents.btnAdAllAddFontColor
                        el.style.fontSize = this.contents.btnAdAllFontSize
                        el.style.borderWidth = this.contents.btnAdAllBorderWidth
                        el.style.borderStyle = this.contents.btnAdAllBorderStyle
                        el.style.borderColor = this.contents.btnAdAllBorderColor
                        el.style.borderRadius = this.contents.btnAdAllBorderRadius
                        el.style.backgroundColor = this.contents.btnAdAllAddBgColor
                    })
                    document.querySelectorAll('.form-content .ad .el-button--danger').forEach(el => {
                        el.style.height = this.contents.btnAdAllHeight
                        el.style.color = this.contents.btnAdAllDelFontColor
                        el.style.fontSize = this.contents.btnAdAllFontSize
                        el.style.borderWidth = this.contents.btnAdAllBorderWidth
                        el.style.borderStyle = this.contents.btnAdAllBorderStyle
                        el.style.borderColor = this.contents.btnAdAllBorderColor
                        el.style.borderRadius = this.contents.btnAdAllBorderRadius
                        el.style.backgroundColor = this.contents.btnAdAllDelBgColor
                    })
                    document.querySelectorAll('.form-content .ad .el-button--warning').forEach(el => {
                        el.style.height = this.contents.btnAdAllHeight
                        el.style.color = this.contents.btnAdAllWarnFontColor
                        el.style.fontSize = this.contents.btnAdAllFontSize
                        el.style.borderWidth = this.contents.btnAdAllBorderWidth
                        el.style.borderStyle = this.contents.btnAdAllBorderStyle
                        el.style.borderColor = this.contents.btnAdAllBorderColor
                        el.style.borderRadius = this.contents.btnAdAllBorderRadius
                        el.style.backgroundColor = this.contents.btnAdAllWarnBgColor
                    })
                })
            },
            // 表格
            rowStyle({row, rowIndex}) {
                if (rowIndex % 2 == 1) {
                    if (this.contents.tableStripe) {
                        return {color: this.contents.tableStripeFontColor}
                    }
                } else {
                    return ''
                }
            },
            cellStyle({row, rowIndex}) {
                if (rowIndex % 2 == 1) {
                    if (this.contents.tableStripe) {
                        return {backgroundColor: this.contents.tableStripeBgColor}
                    }
                } else {
                    return ''
                }
            },
            headerRowStyle({row, rowIndex}) {
                return {color: this.contents.tableHeaderFontColor}
            },
            headerCellStyle({row, rowIndex}) {
                return {backgroundColor: this.contents.tableHeaderBgColor}
            },
            // 表格按钮
            contentTableBtnStyleChange() {
                // this.$nextTick(()=>{
                //   setTimeout(()=>{
                //     document.querySelectorAll('.table-content .tables .el-table__body .el-button--success').forEach(el=>{
                //       el.style.height = this.contents.tableBtnHeight
                //       el.style.color = this.contents.tableBtnDetailFontColor
                //       el.style.fontSize = this.contents.tableBtnFontSize
                //       el.style.borderWidth = this.contents.tableBtnBorderWidth
                //       el.style.borderStyle = this.contents.tableBtnBorderStyle
                //       el.style.borderColor = this.contents.tableBtnBorderColor
                //       el.style.borderRadius = this.contents.tableBtnBorderRadius
                //       el.style.backgroundColor = this.contents.tableBtnDetailBgColor
                //     })
                //     document.querySelectorAll('.table-content .tables .el-table__body .el-button--primary').forEach(el=>{
                //       el.style.height = this.contents.tableBtnHeight
                //       el.style.color = this.contents.tableBtnEditFontColor
                //       el.style.fontSize = this.contents.tableBtnFontSize
                //       el.style.borderWidth = this.contents.tableBtnBorderWidth
                //       el.style.borderStyle = this.contents.tableBtnBorderStyle
                //       el.style.borderColor = this.contents.tableBtnBorderColor
                //       el.style.borderRadius = this.contents.tableBtnBorderRadius
                //       el.style.backgroundColor = this.contents.tableBtnEditBgColor
                //     })
                //     document.querySelectorAll('.table-content .tables .el-table__body .el-button--danger').forEach(el=>{
                //       el.style.height = this.contents.tableBtnHeight
                //       el.style.color = this.contents.tableBtnDelFontColor
                //       el.style.fontSize = this.contents.tableBtnFontSize
                //       el.style.borderWidth = this.contents.tableBtnBorderWidth
                //       el.style.borderStyle = this.contents.tableBtnBorderStyle
                //       el.style.borderColor = this.contents.tableBtnBorderColor
                //       el.style.borderRadius = this.contents.tableBtnBorderRadius
                //       el.style.backgroundColor = this.contents.tableBtnDelBgColor
                //     })

                //   }, 50)
                // })
            },
            // 分页
            contentPageStyleChange() {
                let arr = []
                if (this.contents.pageTotal) arr.push('total')
                if (this.contents.pageSizes) arr.push('sizes')
                if (this.contents.pagePrevNext) {
                    arr.push('prev')
                    if (this.contents.pagePager) arr.push('pager')
                    arr.push('next')
                }
                if (this.contents.pageJumper) arr.push('jumper')
                this.layouts = arr.join()
                this.contents.pageEachNum = 10
            },

            init() {
            },
            search() {
                this.pageIndex = 1;
                this.getDataList();
            },
            // 获取数据列表
            getDataList() {
                this.dataListLoading = true;
                let params = {
                    page: this.pageIndex,
                    limit: this.pageSize,
                    sort: 'id',
                }

                                                             
                if (this.searchForm.jiaoshiName!= '' && this.searchForm.jiaoshiName!= undefined) {
                    params['jiaoshiName'] = '%' + this.searchForm.jiaoshiName + '%'
                }
                     
                if (this.searchForm.jiaoshiTypes!= '' && this.searchForm.jiaoshiTypes!= undefined) {
                    params['jiaoshiTypes'] = this.searchForm.jiaoshiTypes
                }
                     
                if (this.searchForm.jiaoshiStatusTypes!= '' && this.searchForm.jiaoshiStatusTypes!= undefined) {
                    params['jiaoshiStatusTypes'] = this.searchForm.jiaoshiStatusTypes
                }
                                                                                                                                                                 
                if (this.searchForm.laoshiName!= '' && this.searchForm.laoshiName!= undefined) {
                    params['laoshiName'] = '%' + this.searchForm.laoshiName + '%'
                }
                                                                                                                                             
                if (this.searchForm.yonghuName!= '' && this.searchForm.yonghuName!= undefined) {
                    params['yonghuName'] = '%' + this.searchForm.yonghuName + '%'
                }
                                                                                                                                                                                                        
                params['shujiyuyueDelete'] = 1// 逻辑删除字段 1 未删除 2 删除


                this.$http({
                    url: "shujiyuyue/page",
                    method: "get",
                    params: params
                }).then(({data}) => {
                    if(data && data.code === 0){
                        this.dataList = data.data.list;
                        this.totalPage = data.data.total;
                    }else{
                        this.dataList = [];
                        this.totalPage = 0;
                    }
                    this.dataListLoading = false;
                });

                //查询级联表搜索条件所有列表
                this.$http({
                    url: "dictionary/page?dicCode=jiaoshi_types&page=1&limit=100",
                    method: "get",
                    page: 1,
                    limit: 100,
                }).then(({data}) => {
                    if(data && data.code === 0){
                        this.jiaoshiTypesSelectSearch = data.data.list;
                    }
                });
                this.$http({
                    url: "dictionary/page?dicCode=jiaoshi_status_types&page=1&limit=100",
                    method: "get",
                    page: 1,
                    limit: 100,
                }).then(({data}) => {
                    if(data && data.code === 0){
                        this.jiaoshiStatusTypesSelectSearch = data.data.list;
                    }
                });
                //查询当前表搜索条件所有列表
            },
            //每页数
            sizeChangeHandle(val) {
                this.pageSize = val;
                this.pageIndex = 1;
                this.getDataList();
            },
            // 当前页
            currentChangeHandle(val) {
                this.pageIndex = val;
                this.getDataList();
            },
            // 多选
            selectionChangeHandler(val) {
                this.dataListSelections = val;
            },
            // 添加/修改
            addOrUpdateHandler(id, type) {
                this.showFlag = false;
                this.addOrUpdateFlag = true;
                this.crossAddOrUpdateFlag = false;
                if (type != 'info') {
                    type = 'else';
                }
                this.$nextTick(() => {
                    this.$refs.addOrUpdate.init(id, type);
                });
            },
            // 下载
            download(file) {
                window.open(" ${file} ")
            },
            // 删除
            deleteHandler(id) {
                var ids = id ? [Number(id)] : this.dataListSelections.map(item => {
                    return Number(item.id);
                });

                this.$confirm(`确定进行[${id ? "删除" : "批量删除"}]操作?`, "提示", {
                    confirmButtonText: "确定",
                    cancelButtonText: "取消",
                    type: "warning"
                }).then(() => {
                    this.$http({
                        url: "shujiyuyue/delete",
                        method: "post",
                        data: ids
                    }).then(({data}) => {
                        if(data && data.code === 0){
                            this.$message({
                                message: "操作成功",
                                type: "success",
                                duration: 1500,
                                onClose: () => {
                                    this.search();
                                }
                            });
                        }else{
                            this.$message.error(data.msg);
                        }
                    });
                });
            },
            // 导入功能上传文件成功后调用导入方法
            shujiyuyueUploadSuccess(data){
                let _this = this;
                _this.$http({
                    url: "shujiyuyue/batchInsert?fileName=" + data.file,
                    method: "get"
                }).then(({data}) => {
                    if(data && data.code === 0){
                        _this.$message({
                            message: "导入教室预约数据成功",
                            type: "success",
                            duration: 1500,
                            onClose: () => {
                                _this.search();
                            }
                        });
                    }else{
                        _this.$message.error(data.msg);
                    }
                });

            },
            // 导入功能上传文件失败后调用导入方法
            shujiyuyueUploadError(data){
                this.$message.error('上传失败');
            },
            openYesnoTypes(id,jiaoshiId) {
                let _this = this;
                _this.form.id = null;
                _this.form.jiaoshiId = null;
                _this.form.id = id;
                _this.form.jiaoshiId = jiaoshiId;
                _this.form.shujiyuyueYesnoTypes = "请选择审核结果";
                _this.shujiyuyueYesnoTypesVisible = true;
            },

            shujiyuyueYesnoTypesShenhe() {
                let _this = this;
                if(_this.form.shujiyuyueYesnoTypes == "请选择审核结果"){
                    alert("请选择审核结果");
                    return false;
                }
                _this.$http({
                    url:`shujiyuyue/update`,
                    method: "post",
                    data: _this.form
                }).then(({ data }) => {
                    if (data && data.code === 0) {
                        _this.$message({
                            message: "审核成功",
                            type: "success",
                            duration: 1500,
                            onClose: () => {
                                _this.shujiyuyueYesnoTypesVisible = false;
                                _this.search();
								let data = {
									id:_this.form.jiaoshiId,
									jiaoshiStatusTypes:2
								}
								_this.$http({
								    url: "jiaoshi/update",
								    method: "post",
								    data: data
								}).then(({data}) => {
								    if(data && data.code === 0){
										_this.search();
								    }else{
								        _this.$message.error(data.msg);
								    }
								});
                            }
                        });
                    } else {
                        _this.$message.error(data.msg);
                        _this.shujiyuyueYesnoTypesVisible = false;
                    }
                });
            },        }
    };
</script>
<style lang="scss" scoped>
.slt {
    margin: 0 !important;
    display: flex;
  }

  .ad {
    margin: 0 !important;
    display: flex;
  }

  .pages {
    & /deep/ el-pagination__sizes{
      & /deep/ el-input__inner {
        height: 22px;
        line-height: 22px;
      }
    }
  }
  

  .el-button+.el-button {
    margin:0;
  } 

  .tables {
	& /deep/ .el-button--success {
		height: 40px;
		color: #333;
		font-size: 14px;
		border-width: 1px;
		border-style: solid;
		border-color: #DCDFE6;
		border-radius: 4px;
		background-color: rgba(117, 113, 249, 1);
	}
	
	& /deep/ .el-button--primary {
		height: 40px;
		color: #333;
		font-size: 14px;
		border-width: 1px;
		border-style: solid;
		border-color: #DCDFE6;
		border-radius: 4px;
		background-color: rgba(102, 204, 255, 1);
	}
	
	& /deep/ .el-button--danger {
		height: 40px;
		color: #333;
		font-size: 14px;
		border-width: 1px;
		border-style: solid;
		border-color: #DCDFE6;
		border-radius: 4px;
		background-color: rgba(204, 255, 102, 1);
	}

    & /deep/ .el-button {
      margin: 4px;
    }
  }
</style>


