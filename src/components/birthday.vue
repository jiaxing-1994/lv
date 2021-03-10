<template>
  <section v-show="!isEdit">
    <a-button style="margin-right: 10px" @click="editBtn">编辑</a-button>
    <a-button @click="downloadBtn" type="primary">保存图片</a-button>
  </section>
  <a-button v-show="isEdit" type="primary" @click="isEdit = false">编辑完成</a-button>
  <div id="contain">
    <div id="box"
         :style="backgroundStyle"
    >
      <div class="title"
           :style="titleStyle"
           v-for="(i, index) in titleState"
           :key="i.id"
      >
        <div class="input-box" v-show="isEdit">
          <a-input v-model:value="i.text"></a-input>
          <a-button class="del-btn" shape="circle" @click="delTitleBtn(index)"><DeleteOutlined/></a-button>
        </div>
        <span v-show="!isEdit">{{i.text}}</span>
      </div>
      <a-button type="primary" v-show="isEdit" @click="addTitleBtn">添加大标题</a-button>
      <div class="info"
           :style="infoStyle"
           v-for="(i, index) in infoState"
           :key="i.id"
      >
        <div class="input-box" v-show="isEdit">
          <a-textarea v-model:value="i.text"></a-textarea>
          <a-button class="del-btn" shape="circle" @click="delInfoBtn(index)"><DeleteOutlined/></a-button>
        </div>
        <span v-show="!isEdit">{{i.text}}</span>
      </div>
      <a-button type="primary" v-show="isEdit" @click="addInfoBtn">添加小标题</a-button>
      <div class="table">
        <a-button type="primary" @click="resetTable" v-show="isCreated && isEdit">重新选择文件</a-button>
        <div class="show-table" v-show="isCreated">
          <div class="table-row table-h">
            <div class="table-col">所属机构</div>
            <div class="table-col">姓名</div>
          </div>
          <div class="table-row"
               v-for="(i, index) in tableData"
               :key="index"
          >
            <div class="table-col">{{i.title}}</div>
            <div class="table-col">{{i.name}}</div>
          </div>
        </div>
        <div class="read-table" v-show="isEdit">
          <!--        <div class="file-list">-->
          <!--          <div class="file-item"-->
          <!--            v-for="(i, index) in fileList"-->
          <!--               :key="index"-->
          <!--          >-->
          <!--            <span>{{i.name}}</span>-->
          <!--            <span class="del" @click="delFileBtn(index)"><DeleteOutlined/></span>-->
          <!--          </div>-->
          <!--        </div>-->
          <div class="table-btn">
            <a-upload
              v-model:file-list="fileList"
              name="file"
              :multiple="false"
              :before-upload="beforeUpload"
              :showUploadList="false"
            >
              <a-button type="primary" @click="openModel">
                点击选择excel文件生成表格
              </a-button>
            </a-upload>
            <!--          <a-button type="primary" @click="createTable">生成表格</a-button>-->
          </div>
        </div>
      </div>
      <div class="footer"
           :style="footerStyle"
      >
        <h3>人力资源部</h3>
        <h3>成都网阔信息技术股份有限公司</h3>
        <h3>{{nowDate}}</h3>
      </div>
    </div>
    <div>
      <div id="edit">
        <a-form>
          <div class="need-box">
            <h3>这三个必须填！(选择文件前，先把这三个表头填好)</h3>
            <a-form-item label="姓名列的表头">
              <a-input v-model:value="fileTitle.name"></a-input>
            </a-form-item>
            <a-form-item label="出生日期/身份证列的表头">
              <a-input v-model:value="fileTitle.date"></a-input>
            </a-form-item>
            <a-form-item label="机构组织列的表头">
              <a-input v-model:value="fileTitle.title"></a-input>
            </a-form-item>
          </div>
          <a-form-item label="背景图">
            <div class="background-btn">
              <a-upload
                v-model:file-list="backgroundFile"
                name="file"
                :multiple="false"
                :before-upload="backgroundBeforeUpload"
                :showUploadList="false"
              >
                <a-button>
                  {{backgroundFile.length === 0 ? '选择图片' : '重新选择'}}
                </a-button>
              </a-upload>
              <a-button class="del-background-btn" v-show="backgroundFile.length > 0" @click="delBackgroundBtn">删除背景图</a-button>
            </div>
          </a-form-item>
          <a-form-item label="大标题字体样式">
            <a-select
              v-model:value="config.currTitleFontFamily"
            >
              <a-select-option
                v-for="(i, index) in config.fontFamilyList"
                :key="index"
                :value="i"
              >{{i}}</a-select-option>
            </a-select>
          </a-form-item>
          <a-form-item label="小标题字体样式">
            <a-select
              v-model:value="config.currInfoFontFamily"
            >
              <a-select-option
                v-for="(i, index) in config.fontFamilyList"
                :key="index"
                :value="i"
              >{{i}}</a-select-option>
            </a-select>
          </a-form-item>
          <a-form-item label="底部文字字体样式">
            <a-select
              v-model:value="config.currFooterFontFamily"
            >
              <a-select-option
                v-for="(i, index) in config.fontFamilyList"
                :key="index"
                :value="i"
              >{{i}}</a-select-option>
            </a-select>
          </a-form-item>
        </a-form>
      </div>
    </div>
  </div>
</template>

<script>
import { ref, reactive, computed } from 'vue';
import XLSX from 'xlsx';
import moment from 'moment';
import h2c from 'html2canvas';
import { saveAs } from 'file-saver';
import { fonts } from '../utils';
import {
  DeleteOutlined,
} from '@ant-design/icons-vue';
export default {
  name: '',
  components: {
    DeleteOutlined,
  },
  setup() {
    const isEdit = ref(false);
    const { nowDate, editBtn, downloadBtn } = useBtn(isEdit);
    const { titleState, addTitleBtn, delTitleBtn } = useTitle();
    const { infoState, addInfoBtn, delInfoBtn } = useInfo();
    const { fileList, tableData, isCreated, fileTitle, beforeUpload, createTable, delFileBtn, resetTable } = useTable();

    const { config, titleStyle, infoStyle, footerStyle, backgroundFile, backgroundStyle, backgroundBeforeUpload, delBackgroundBtn } = useConfig();
    return {
      nowDate,
      config,
      titleStyle,
      infoStyle,
      footerStyle,
      isEdit,
      isCreated,
      titleState,
      infoState,
      fileList,
      tableData,
      backgroundFile,
      backgroundStyle,
      fileTitle,
      editBtn,
      downloadBtn,
      addTitleBtn,
      delTitleBtn,
      addInfoBtn,
      delInfoBtn,
      beforeUpload,
      createTable,
      delFileBtn,
      backgroundBeforeUpload,
      delBackgroundBtn,
      resetTable,
    };
  },
}
// 大标题
function useTitle() {
  const titleState = ref([{
    id: 1,
    text: moment(new Date()).format('YYYY-MM-DD'),
  }, {
    id: 2,
    text: '生日福利',
  }]);
  const addTitleBtn = () => {
    titleState.value.push({
      id: titleState.value.length + 1,
      text: ref(''),
    })
  }
  const delTitleBtn = (index) => {
    titleState.value.splice(index, 1);
    titleState.value = titleState.value.map((item, index) => ({
      ...item,
      id: index,
    }));
  };
  return {
    titleState,
    addTitleBtn,
    delTitleBtn,
  }
}
// 小标题
function useInfo() {
  const infoState = ref([{
    id: 1,
    text: '各位伙伴:',
  }, {
    id: 2,
    text: '大家好，春去秋来，暖冬将至。在这里祝10月的生日伙伴，生日快乐！',
  }, {
    id: 3,
    text: '10月的生日福利为蛋糕卡，片区的伙伴将邮寄生日礼物，福州、昆明的伙伴将由子公司自行安排福利~',
  }, {
    id: 4,
    text: '请成都网阔的生日伙伴前往人力资源部吕佳芮处领取蛋糕卡，谢谢。',
  }]);
  const addInfoBtn = () => {
    infoState.value.push({
      id: infoState.value.length + 1,
      text: ref(''),
    })
  }
  const delInfoBtn = (index) => {
    infoState.value.splice(index, 1);
    infoState.value = infoState.value.map((item, index) => ({
      ...item,
      id: index,
    }));
  };
  return {
    infoState,
    addInfoBtn,
    delInfoBtn,
  }
}
// 通用按钮
function useBtn(isEdit) {
  const nowDate = moment(new Date()).format('YYYY年MM月DD日');
  const editBtn = () => {
    isEdit.value =! isEdit.value;
  }
  const downloadBtn = () => {
    h2c(document.getElementById('box')).then((canvas => {
      canvas.toBlob((blob) => {
        saveAs(blob, '生日');
      });
    }));
  }
  return {
    nowDate,
    editBtn,
    downloadBtn,
  }
}
// 表格文件读取
function useTable() {
  const rules = {
    name: [
      { require: true, message: '请输入姓名列的表头' },
    ],
    title: [
      { require: true, message: '请输入所属机构列的表头' },
    ],
    date: [
      { require: true, message: '请输入出生日期列的表头(身份证的表头也可以哒)' },
    ]
  };
  const fileFormState = reactive({
    name: '',
    title: '',
    date: '',
  });

  const modelFormRef = ref();
  const month = ref(3);
  const tableData = ref([]);
  const fileList = ref([]);
  const isCreated = ref(false);
  const fileTitle = reactive({
    name: '姓名',
    title: '机构组织',
    date: '出生日期'
  });
  const isVisble = ref(false);
  let workbooks = [];
  // 自动生成表格
  const beforeUpload = (file) => {
    readFile(file).then((workbook) => {
      getTableData(workbook, fileTitle.name, fileTitle.date, fileTitle.title);
      if (tableData.value.length > 0) {
        isCreated.value = true;
      }
    });
  }
  // 手动生成表格
  const createTable = () => {
    const promises = [];
    fileList.value.forEach((fileItem) => {
      promises.push(readFile(fileItem.originFileObj));
    });
    Promise.all(promises).then((res) => {
      workbooks = res;
      workbooks.forEach((workbook) => {
        getTableData(workbook, fileTitle.name, fileTitle.date, fileTitle.title);
      });
      if (tableData.value.length > 0) {
        isCreated.value = true;
      }
    });
  }
  // 读取文件
  const readFile = (file) => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const workbook = XLSX.read(e.target.result, { type: 'binary', cellDates: true });
        resolve(workbook);
      };
      reader.readAsBinaryString(file);
    });
  }
  // 获取表格数据
  const getTableData = (workbook, name, date, title) => {
    const result = [];
    workbook.SheetNames.forEach((sheetName) => {
      const currSheet = workbook.Sheets[sheetName];
      const jsonSheet = XLSX.utils.sheet_to_json(currSheet);
      jsonSheet.forEach((item) => {
        for (const key in item) {
          if (key === date) {
            if (handleCurrentDay(item[key])) {
              result.push({
                name: item[name],
                title: item[title],
              });
            }
          }
        }
      })
    });
    getNoRepeatData(result);
  };
  // 去重
  const getNoRepeatData = (newData) => {
    const setData = new Set();
    tableData.value.forEach((item) => {
      setData.add([item.name, item.title]);
    });
    newData.forEach((item) => {
      setData.add([item.name, item.title]);
    });
    const result = [];
    setData.forEach((item) => {
      result.push({
        name: item[0],
        title: item[1],
      });
    });
    tableData.value = result;
  }
  // 判断是否正确月份
  const handleCurrentDay = (value) => {
    const idReg = /^[1-9]\d{7}((0\d)|(1[0-2]))(([0|1|2]\d)|3[0-1])\d{3}$|^[1-9]\d{5}[1-9]\d{3}((0\d)|(1[0-2]))(([0|1|2]\d)|3[0-1])\d{3}([0-9]|X)$/;
    if (value instanceof Date) {
      // 说明是生日;
      console.log(typeof moment(value).format('MM'));
      console.log(typeof month.value);
      if (Number(moment(value).format('MM')) === Number(month.value)) {
        return true;
      }
    }
    if (idReg.test(value)) {
      // 说明是身份证
      let birthday = '';
      if(value != null && value != ""){
        if(value.length == 15){
          birthday = '19' + value.substr(6,6);
        } else if(value.length == 18){
          birthday = value.substr(6,8);
        }
        birthday = birthday.replace(/(.{4})(.{2})/,"$1-$2-");
        if (Number(moment(birthday).format('MM')) === Number(month.value)) {
          return true;
        }
      }
    }
    return false;
  }
  // 删除文件
  const delFileBtn = (index) => {
    fileList.value.splice(index, 1);
  }

  // 打开信息框
  const openModel = () => {
    isVisble.value = !isVisble.value;
  }
  // 文件确认框
  const fileComfirm = () => {
    modelFormRef.value.validate().then(() => {
      console.log(fileFormState);
    });
  }
  // 重置表格
  const resetTable = () => {
    fileList.value = [];
    tableData.value = [];
    workbooks = [];
    isCreated.value = false;
  }
  return {
    fileTitle,
    isCreated,
    isVisble,
    fileList,
    tableData,
    rules,
    modelFormRef,
    fileFormState,
    openModel,
    beforeUpload,
    createTable,
    delFileBtn,
    resetTable,
    fileComfirm,
  }
}

// 配置相关
function useConfig() {
  const config = reactive({
    currTitleFontFamily: fonts[0],
    currFooterFontFamily: fonts[0],
    currInfoFontFamily: fonts[0],
    fontFamilyList: fonts,
    backgroundImage: '',
  })
  const backgroundFile = ref([]);
  const backgroundBeforeUpload = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      config.backgroundImage = e.target.result;
    };
    reader.readAsDataURL(file);
  }
  const delBackgroundBtn = () => {
    config.backgroundImage = '';
    backgroundFile.value = [];
  };
  const backgroundStyle = computed(() => {
    return {
      backgroundImage: `url(${config.backgroundImage}`,
      backgroundRepeat: 'repeat',
      backgroundPosition: 'center',
      backgroundSize: '100%',
    }
  });
  const titleStyle = computed(() => {
    return {
      fontFamily: config.currTitleFontFamily,
    }
  });
  const infoStyle = computed(() => {
    return {
      fontFamily: config.currInfoFontFamily,
    }
  });
  const footerStyle = computed(() => {
    return {
      fontFamily: config.currFooterFontFamily,
    }
  });
  return {
    config,
    titleStyle,
    infoStyle,
    footerStyle,
    backgroundStyle,
    backgroundFile,
    backgroundBeforeUpload,
    delBackgroundBtn,
  }
}
</script>

<style lang="less" scoped>
@table-border-color: #333;
#contain{
  width: 100%;
  display: flex;
  justify-content: flex-start;
}
#box{
  min-width: 400px;
  max-width: 450px;
  padding: 10px;
  margin: 10px;
  border: solid 4px #ddd;
  .title{
    font-size: 30px;
    font-weight: bold;
  }
  .info{
    width: 100%;
    padding: 0 40px;
    font-size: 14px;
    text-align: left;
    text-indent: 2em;
    background: #ffffff9e;
  }
  .info:first-child{
    text-indent: 0em;
  }
  .del-btn{
    margin-left: 10px;
  }
  .table{
    width: 80%;
    margin: 10px auto;
    background: #fff;
    .show-table{
      border-top: solid 1px @table-border-color;
      border-right: solid 1px @table-border-color;
      .table-row{
        display: flex;
        .table-col{
          flex: 1;
          border-bottom: solid 1px @table-border-color;
          border-left: solid 1px @table-border-color;
        }
      }
      .table-h{
        background: #ddd;
        .table-col{
          font-weight: bold;
        }
      }
    }
    .read-table{
      .file-list{
        .file-item{
          color: dodgerblue;
          .del{
            color: red;
            margin-left: 20px;
          }
        }
      }
      .table-btn{
      }
    }
  }
  .footer{
    h3{
      font-size: 14px;
      margin: 0;
    }
  }
  .input-box{
    display: flex;
  }
}
.background-btn{
  display: flex;
  align-items: center;
  .del-background-btn{
    margin-left: 10px;
    color: red;
  }
}
:deep(.ant-form-item){
  display: flex;
  justify-content: flex-start;
  .ant-form-item-label{
    width: 30%;
    display: flex;
    align-items: center;
    padding: 0;
  }
  .ant-form-item-control-wrapper{
    width: 50%;
  }
}
.ant-input{
  margin-bottom: 10px;
}
#edit{
  width: 500px;
  padding: 30px;
  .need-box{
    background: #efefef;
    padding: 5px;
  }
}
</style>
