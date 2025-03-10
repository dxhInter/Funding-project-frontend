<template>
  <div class="app-container">
    <el-card>
      <div slot="header" class="clearfix">
        <span>预算报表系统</span>
        <div class="operations">
          <el-button type="primary" icon="el-icon-download" size="mini" @click="exportExcel">导出Excel</el-button>
          <el-button type="success" icon="el-icon-plus" size="mini" @click="handleAdd">新增</el-button>
        </div>
      </div>

      <!-- 表格部分 -->
      <el-table
        :data="tableData"
        border
        style="width: 100%"
        :span-method="objectSpanMethod"
        :header-cell-style="headerCellStyle"
        :cell-style="cellStyle"
      >
        <!-- 预算科目名称列 -->
        <el-table-column prop="name" label="预算科目名称" width="220">
          <template slot-scope="scope">
            <span :class="{ 'font-bold': scope.row.isTotal || scope.row.isSubtotal }">{{ scope.row.name }}</span>
          </template>
        </el-table-column>

        <!-- 第一年 -->
        <el-table-column label="第一年" align="center">
          <el-table-column prop="year1.fiscal" label="财政经费" width="120" align="right">
            <template slot-scope="scope">
              <span>{{ formatNumber(scope.row.year1.fiscal) }}</span>
            </template>
          </el-table-column>
          <el-table-column prop="year1.selfRaised" label="自筹经费" width="120" align="right">
            <template slot-scope="scope">
              <span>{{ formatNumber(scope.row.year1.selfRaised) }}</span>
            </template>
          </el-table-column>
          <el-table-column prop="year1.outsourcing" label="外委金额" width="120" align="right">
            <template slot-scope="scope">
              <span>{{ formatNumber(scope.row.year1.outsourcing) }}</span>
            </template>
          </el-table-column>
          <el-table-column prop="year1.total" label="小计" width="120" align="right">
            <template slot-scope="scope">
              <span :class="{ 'font-bold': true }">{{ formatNumber(scope.row.year1.total) }}</span>
            </template>
          </el-table-column>
        </el-table-column>

        <!-- 第二年 -->
        <el-table-column label="第二年" align="center">
          <el-table-column prop="year2.fiscal" label="财政经费" width="120" align="right">
            <template slot-scope="scope">
              <span>{{ formatNumber(scope.row.year2.fiscal) }}</span>
            </template>
          </el-table-column>
          <el-table-column prop="year2.selfRaised" label="自筹经费" width="120" align="right">
            <template slot-scope="scope">
              <span>{{ formatNumber(scope.row.year2.selfRaised) }}</span>
            </template>
          </el-table-column>
          <el-table-column prop="year2.outsourcing" label="外委金额" width="120" align="right">
            <template slot-scope="scope">
              <span>{{ formatNumber(scope.row.year2.outsourcing) }}</span>
            </template>
          </el-table-column>
          <el-table-column prop="year2.total" label="小计" width="120" align="right">
            <template slot-scope="scope">
              <span :class="{ 'font-bold': true }">{{ formatNumber(scope.row.year2.total) }}</span>
            </template>
          </el-table-column>
        </el-table-column>

        <!-- 第三年 -->
        <el-table-column label="第三年" align="center">
          <el-table-column prop="year3.fiscal" label="财政经费" width="120" align="right">
            <template slot-scope="scope">
              <span>{{ formatNumber(scope.row.year3.fiscal) }}</span>
            </template>
          </el-table-column>
          <el-table-column prop="year3.selfRaised" label="自筹经费" width="120" align="right">
            <template slot-scope="scope">
              <span>{{ formatNumber(scope.row.year3.selfRaised) }}</span>
            </template>
          </el-table-column>
          <el-table-column prop="year3.outsourcing" label="外委金额" width="120" align="right">
            <template slot-scope="scope">
              <span>{{ formatNumber(scope.row.year3.outsourcing) }}</span>
            </template>
          </el-table-column>
          <el-table-column prop="year3.total" label="小计" width="120" align="right">
            <template slot-scope="scope">
              <span :class="{ 'font-bold': true }">{{ formatNumber(scope.row.year3.total) }}</span>
            </template>
          </el-table-column>
        </el-table-column>

        <!-- 总计 -->
        <el-table-column label="总计" align="center">
          <el-table-column prop="total.fiscal" label="财政经费" width="120" align="right">
            <template slot-scope="scope">
              <span :class="{ 'font-bold': scope.row.isTotal || scope.row.isSubtotal }">{{ formatNumber(scope.row.total.fiscal) }}</span>
            </template>
          </el-table-column>
          <el-table-column prop="total.selfRaised" label="自筹经费" width="120" align="right">
            <template slot-scope="scope">
              <span :class="{ 'font-bold': scope.row.isTotal || scope.row.isSubtotal }">{{ formatNumber(scope.row.total.selfRaised) }}</span>
            </template>
          </el-table-column>
          <el-table-column prop="total.outsourcing" label="外委金额" width="120" align="right">
            <template slot-scope="scope">
              <span :class="{ 'font-bold': scope.row.isTotal || scope.row.isSubtotal }">{{ formatNumber(scope.row.total.outsourcing) }}</span>
            </template>
          </el-table-column>
          <el-table-column prop="total.total" label="小计" width="120" align="right">
            <template slot-scope="scope">
              <span :class="{ 'font-bold': scope.row.isTotal || scope.row.isSubtotal }">{{ formatNumber(scope.row.total.total) }}</span>
            </template>
          </el-table-column>
        </el-table-column>

        <!-- 操作列 -->
        <el-table-column label="操作" width="120" fixed="right" v-if="!onlyDisplay">
          <template slot-scope="scope">
            <el-button
              v-if="!scope.row.isTotal && !scope.row.isSubtotal && !scope.row.isHeader"
              type="text"
              size="mini"
              icon="el-icon-edit"
              @click="handleEdit(scope.row)"
              class="edit-button"
            >编辑</el-button>
          </template>
        </el-table-column>
      </el-table>
    </el-card>

    <!-- 编辑/添加表单对话框 -->
    <el-dialog :title="dialogTitle" :visible.sync="dialogVisible" width="700px" append-to-body>
      <el-form ref="form" :model="form" :rules="rules" label-position="top">
        <el-form-item label="项目名称" prop="name">
          <el-input v-model="form.name" placeholder="请输入项目名称" :disabled="form.isHeader || form.isTotal || form.isSubtotal"/>
        </el-form-item>

        <el-form-item label="是否为标题" v-if="isAdding">
          <el-switch v-model="form.isHeader"></el-switch>
        </el-form-item>

        <template v-if="!form.isHeader">
          <el-divider content-position="left">第一年</el-divider>
          <el-row :gutter="20">
            <el-col :span="8">
              <el-form-item label="财政经费">
                <el-input-number v-model="form.year1.fiscal" :precision="2" :step="1000" :min="0" style="width: 100%"></el-input-number>
              </el-form-item>
            </el-col>
            <el-col :span="8">
              <el-form-item label="自筹经费">
                <el-input-number v-model="form.year1.selfRaised" :precision="2" :step="1000" :min="0" style="width: 100%"></el-input-number>
              </el-form-item>
            </el-col>
            <el-col :span="8">
              <el-form-item label="外委金额">
                <el-input-number v-model="form.year1.outsourcing" :precision="2" :step="1000" :min="0" style="width: 100%"></el-input-number>
              </el-form-item>
            </el-col>
          </el-row>

          <el-divider content-position="left">第二年</el-divider>
          <el-row :gutter="20">
            <el-col :span="8">
              <el-form-item label="财政经费">
                <el-input-number v-model="form.year2.fiscal" :precision="2" :step="1000" :min="0" style="width: 100%"></el-input-number>
              </el-form-item>
            </el-col>
            <el-col :span="8">
              <el-form-item label="自筹经费">
                <el-input-number v-model="form.year2.selfRaised" :precision="2" :step="1000" :min="0" style="width: 100%"></el-input-number>
              </el-form-item>
            </el-col>
            <el-col :span="8">
              <el-form-item label="外委金额">
                <el-input-number v-model="form.year2.outsourcing" :precision="2" :step="1000" :min="0" style="width: 100%"></el-input-number>
              </el-form-item>
            </el-col>
          </el-row>

          <el-divider content-position="left">第三年</el-divider>
          <el-row :gutter="20">
            <el-col :span="8">
              <el-form-item label="财政经费">
                <el-input-number v-model="form.year3.fiscal" :precision="2" :step="1000" :min="0" style="width: 100%"></el-input-number>
              </el-form-item>
            </el-col>
            <el-col :span="8">
              <el-form-item label="自筹经费">
                <el-input-number v-model="form.year3.selfRaised" :precision="2" :step="1000" :min="0" style="width: 100%"></el-input-number>
              </el-form-item>
            </el-col>
            <el-col :span="8">
              <el-form-item label="外委金额">
                <el-input-number v-model="form.year3.outsourcing" :precision="2" :step="1000" :min="0" style="width: 100%"></el-input-number>
              </el-form-item>
            </el-col>
          </el-row>
        </template>
      </el-form>
      <div slot="footer" class="dialog-footer">
        <el-button @click="dialogVisible = false">取 消</el-button>
        <el-button type="primary" @click="submitForm">确 定</el-button>
      </div>
    </el-dialog>
  </div>
</template>

<script>
import { deepClone } from '@/utils/index';
import { saveAs } from 'file-saver';
import XLSX from 'xlsx';

export default {
  name: "BudgetReport",
  props: {
    projectId: {
      type: [String, Number],
      default: null
    },
    onlyDisplay: {
      type: Boolean,
      default: false
    }
  },
  data() {
    return {
      tableData: [],
      dialogVisible: false,
      dialogTitle: '',
      isAdding: false,
      editIndex: -1,
      form: this.getEmptyForm(),
      rules: {
        name: [
          { required: true, message: '请输入项目名称', trigger: 'blur' }
        ]
      },
      // 默认的表格结构
      defaultTableStructure: [
        { name: '总    计', type: 'total', isTotal: true },
        { name: '一、直接费用合计：', type: 'directTotal', isSubtotal: true },
        { name: '（一）费用性支出合计', type: 'expenseTotal', isSubtotal: true },
        { name: '1.人工费', type: 'laborCost' },
        { name: '2.材料费', type: 'materialCost' },
        { name: '其中：低值易耗品', type: 'lowValueConsumables', indent: true },
        { name: '3.仪器设备维护使用租赁费', type: 'equipmentMaintenanceCost' },
        { name: '4.测试化验加工建造费', type: 'testProcessingCost' },
        { name: '5.燃料动力费', type: 'fuelPowerCost' },
        { name: '6.会议费', type: 'meetingCost' },
        { name: '7.差旅费', type: 'travelCost' },
        { name: '8.培训费', type: 'trainingCost' },
        { name: '9.国际合作与交流费', type: 'internationalCooperationCost' },
        { name: '10.场地租赁费', type: 'venueCost' },
        { name: '11.出版/文献/信息传播/知识产权事务费', type: 'publicationCost' },
        { name: '12.劳务费', type: 'laborServiceCost' },
        { name: '13.专家咨询费', type: 'consultationCost' },
        { name: '14.其他支出', type: 'otherExpenses' },
        { name: '（二）资本性支出', type: 'capitalTotal', isSubtotal: true },
        { name: '1.设备购置费', type: 'equipmentPurchaseCost' },
        { name: '2.设备试制费', type: 'equipmentTrialCost' },
        { name: '3.软件开发购置费', type: 'softwareDevelopmentCost' },
        { name: '4.其他资本性支出', type: 'otherCapitalCost' },
        { name: '二、间接费用合计：', type: 'indirectTotal', isSubtotal: true },
        { name: '1.仪器设备房屋使用或折旧', type: 'equipmentDepreciationCost' },
        { name: '2.水、电、气、暖', type: 'utilityCost' },
        { name: '3.有关管理费用', type: 'managementCost' },
        { name: '4.绩效支出', type: 'performanceCost' }
      ]
    };
  },
  created() {
    this.initTableData();
    if (this.projectId) {
      this.fetchData();
    }
  },
  computed: {
    // 计算项目总数（用于生成新项目ID）
    totalItems() {
      return this.tableData.length;
    }
  },
  methods: {
    // 初始化表格数据结构
    initTableData() {
      this.tableData = this.defaultTableStructure.map(item => {
        const row = {
          ...item,
          id: this.generateId(),
          year1: {
            fiscal: 0,
            selfRaised: 0,
            outsourcing: 0,
            total: 0
          },
          year2: {
            fiscal: 0,
            selfRaised: 0,
            outsourcing: 0,
            total: 0
          },
          year3: {
            fiscal: 0,
            selfRaised: 0,
            outsourcing: 0,
            total: 0
          },
          total: {
            fiscal: 0,
            selfRaised: 0,
            outsourcing: 0,
            total: 0
          }
        };
        return row;
      });
    },

    // 从服务器获取数据
    fetchData() {
      // 使用若依框架的请求方式从后端获取数据
      this.$modal.loading("正在加载数据，请稍候...");
      this.$http.get(`/budget/report/${this.projectId}`).then(response => {
        if (response.data.code === 200) {
          const data = response.data.data;
          if (data && data.length > 0) {
            this.tableData = data;
          }
          this.recalculateTotals();
        } else {
          this.$message.error(response.data.msg || '获取数据失败');
        }
        this.$modal.closeLoading();
      }).catch(error => {
        console.error('获取预算报表数据失败', error);
        this.$message.error('获取数据失败，请稍后重试');
        this.$modal.closeLoading();
      });
    },

    // 生成唯一ID
    generateId() {
      return 'budget_' + Date.now() + '_' + Math.floor(Math.random() * 1000);
    },

    // 表格单元格合并方法
    objectSpanMethod({ row, column, rowIndex, columnIndex }) {
      // 可以根据需要实现合并单元格的逻辑
      return {
        rowspan: 1,
        colspan: 1
      };
    },

    // 表头单元格样式
    headerCellStyle() {
      return {
        backgroundColor: '#f5f7fa',
        color: '#606266',
        fontWeight: 'bold',
        textAlign: 'center'
      };
    },

    // 单元格样式
    cellStyle({ row, column, rowIndex, columnIndex }) {
      if (row.isTotal || row.isSubtotal) {
        return {
          backgroundColor: '#f5f7fa',
          fontWeight: 'bold'
        };
      }
      if (row.indent) {
        return {
          paddingLeft: '30px',
        };
      }
      return {};
    },

    // 格式化数字（添加千位分隔符）
    formatNumber(num) {
      if (num === undefined || num === null) return '0.00';
      return num.toLocaleString('zh-CN', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      });
    },

    // 重新计算各个小计和总计
    recalculateTotals() {
      // 先计算每行的总计
      this.tableData.forEach(row => {
        if (row.isHeader) return; // 跳过标题行

        // 计算每年的小计
        row.year1.total = row.year1.fiscal + row.year1.selfRaised + row.year1.outsourcing;
        row.year2.total = row.year2.fiscal + row.year2.selfRaised + row.year2.outsourcing;
        row.year3.total = row.year3.fiscal + row.year3.selfRaised + row.year3.outsourcing;

        // 计算各项的总计
        row.total.fiscal = row.year1.fiscal + row.year2.fiscal + row.year3.fiscal;
        row.total.selfRaised = row.year1.selfRaised + row.year2.selfRaised + row.year3.selfRaised;
        row.total.outsourcing = row.year1.outsourcing + row.year2.outsourcing + row.year3.outsourcing;
        row.total.total = row.total.fiscal + row.total.selfRaised + row.total.outsourcing;
      });

      // 计算各类别小计
      this.calculateSectionTotals('expenseTotal', 1, 14); // 费用性支出合计
      this.calculateSectionTotals('capitalTotal', 15, 18); // 资本性支出
      this.calculateSectionTotals('directTotal', 2, 18); // 直接费用合计
      this.calculateSectionTotals('indirectTotal', 20, 23); // 间接费用合计
      this.calculateSectionTotals('total', 1, 23); // 总计
    },

    // 计算各节的小计
    calculateSectionTotals(sectionType, startIndex, endIndex) {
      const sectionRow = this.tableData.find(row => row.type === sectionType);
      if (!sectionRow) return;

      // 初始化所有值为0
      sectionRow.year1.fiscal = 0;
      sectionRow.year1.selfRaised = 0;
      sectionRow.year1.outsourcing = 0;
      sectionRow.year1.total = 0;

      sectionRow.year2.fiscal = 0;
      sectionRow.year2.selfRaised = 0;
      sectionRow.year2.outsourcing = 0;
      sectionRow.year2.total = 0;

      sectionRow.year3.fiscal = 0;
      sectionRow.year3.selfRaised = 0;
      sectionRow.year3.outsourcing = 0;
      sectionRow.year3.total = 0;

      sectionRow.total.fiscal = 0;
      sectionRow.total.selfRaised = 0;
      sectionRow.total.outsourcing = 0;
      sectionRow.total.total = 0;

      // 汇总指定范围内的所有值
      for (let i = startIndex; i <= endIndex; i++) {
        const row = this.tableData[i];
        if (!row || row.isHeader || row.isSubtotal) continue; // 跳过标题行和小计行

        sectionRow.year1.fiscal += row.year1.fiscal || 0;
        sectionRow.year1.selfRaised += row.year1.selfRaised || 0;
        sectionRow.year1.outsourcing += row.year1.outsourcing || 0;

        sectionRow.year2.fiscal += row.year2.fiscal || 0;
        sectionRow.year2.selfRaised += row.year2.selfRaised || 0;
        sectionRow.year2.outsourcing += row.year2.outsourcing || 0;

        sectionRow.year3.fiscal += row.year3.fiscal || 0;
        sectionRow.year3.selfRaised += row.year3.selfRaised || 0;
        sectionRow.year3.outsourcing += row.year3.outsourcing || 0;
      }

      // 计算每年的小计
      sectionRow.year1.total = sectionRow.year1.fiscal + sectionRow.year1.selfRaised + sectionRow.year1.outsourcing;
      sectionRow.year2.total = sectionRow.year2.fiscal + sectionRow.year2.selfRaised + sectionRow.year2.outsourcing;
      sectionRow.year3.total = sectionRow.year3.fiscal + sectionRow.year3.selfRaised + sectionRow.year3.outsourcing;

      // 计算总计行
      sectionRow.total.fiscal = sectionRow.year1.fiscal + sectionRow.year2.fiscal + sectionRow.year3.fiscal;
      sectionRow.total.selfRaised = sectionRow.year1.selfRaised + sectionRow.year2.selfRaised + sectionRow.year3.selfRaised;
      sectionRow.total.outsourcing = sectionRow.year1.outsourcing + sectionRow.year2.outsourcing + sectionRow.year3.outsourcing;
      sectionRow.total.total = sectionRow.total.fiscal + sectionRow.total.selfRaised + sectionRow.total.outsourcing;
    },

    // 处理添加按钮点击
    handleAdd() {
      this.dialogTitle = '新增预算项';
      this.form = this.getEmptyForm();
      this.isAdding = true;
      this.editIndex = -1;
      this.dialogVisible = true;
    },

    // 处理编辑按钮点击
    handleEdit(row) {
      this.dialogTitle = '编辑预算项';
      this.form = deepClone(row);
      this.isAdding = false;
      this.editIndex = this.tableData.findIndex(item => item.id === row.id);
      this.dialogVisible = true;
    },

    // 获取空的表单对象
    getEmptyForm() {
      return {
        id: this.generateId(),
        name: '',
        isHeader: false,
        isTotal: false,
        isSubtotal: false,
        year1: {
          fiscal: 0,
          selfRaised: 0,
          outsourcing: 0,
          total: 0
        },
        year2: {
          fiscal: 0,
          selfRaised: 0,
          outsourcing: 0,
          total: 0
        },
        year3: {
          fiscal: 0,
          selfRaised: 0,
          outsourcing: 0,
          total: 0
        },
        total: {
          fiscal: 0,
          selfRaised: 0,
          outsourcing: 0,
          total: 0
        }
      };
    },

    // 提交表单
    submitForm() {
      this.$refs.form.validate((valid) => {
        if (valid) {
          // 计算小计
          if (!this.form.isHeader) {
            this.form.year1.total = this.form.year1.fiscal + this.form.year1.selfRaised + this.form.year1.outsourcing;
            this.form.year2.total = this.form.year2.fiscal + this.form.year2.selfRaised + this.form.year2.outsourcing;
            this.form.year3.total = this.form.year3.fiscal + this.form.year3.selfRaised + this.form.year3.outsourcing;

            this.form.total.fiscal = this.form.year1.fiscal + this.form.year2.fiscal + this.form.year3.fiscal;
            this.form.total.selfRaised = this.form.year1.selfRaised + this.form.year2.selfRaised + this.form.year3.selfRaised;
            this.form.total.outsourcing = this.form.year1.outsourcing + this.form.year2.outsourcing + this.form.year3.outsourcing;
            this.form.total.total = this.form.total.fiscal + this.form.total.selfRaised + this.form.total.outsourcing;
          }

          if (this.isAdding) {
            // 如果是标题行，将所有数据置为0
            if (this.form.isHeader) {
              this.resetFormValues();
            }
            // 添加新数据
            this.tableData.push(deepClone(this.form));
          } else {
            // 更新数据
            this.$set(this.tableData, this.editIndex, deepClone(this.form));
          }

          this.recalculateTotals();
          this.dialogVisible = false;

          // 如果有项目ID，则保存到服务器
          if (this.projectId) {
            this.saveToServer();
          }
        } else {
          return false;
        }
      });
    },

    // 将标题行的所有数值重置为0
    resetFormValues() {
      this.form.year1 = {
        fiscal: 0,
        selfRaised: 0,
        outsourcing: 0,
        total: 0
      };
      this.form.year2 = {
        fiscal: 0,
        selfRaised: 0,
        outsourcing: 0,
        total: 0
      };
      this.form.year3 = {
        fiscal: 0,
        selfRaised: 0,
        outsourcing: 0,
        total: 0
      };
      this.form.total = {
        fiscal: 0,
        selfRaised: 0,
        outsourcing: 0,
        total: 0
      };
    },

    // 保存数据到服务器
    saveToServer() {
      const data = {
        projectId: this.projectId,
        budgetItems: this.tableData
      };

      this.$modal.loading("正在保存数据，请稍候...");
      this.$http.post('/budget/report/save', data).then(response => {
        if (response.data.code === 200) {
          this.$message.success('保存成功');
        } else {
          this.$message.error(response.data.msg || '保存失败');
        }
        this.$modal.closeLoading();
      }).catch(error => {
        console.error('保存预算报表数据失败', error);
        this.$message.error('保存失败，请稍后重试');
        this.$modal.closeLoading();
      });
    },

    // 导出Excel
    exportExcel() {
      // 创建工作簿和工作表
      const wb = XLSX.utils.book_new();

      // 准备表头数据
      const headers = [
        ['预算科目名称', '第一年', '', '', '', '第二年', '', '', '', '第三年', '', '', '', '总计', '', '', ''],
        ['', '财政经费', '自筹经费', '外委金额', '小计', '财政经费', '自筹经费', '外委金额', '小计', '财政经费', '自筹经费', '外委金额', '小计', '财政经费', '自筹经费', '外委金额', '小计']
      ];

      // 准备数据
      const data = this.tableData.map(row => {
        return [
          row.name,
          // 第一年
          row.year1.fiscal,
          row.year1.selfRaised,
          row.year1.outsourcing,
          row.year1.total,
          // 第二年
          row.year2.fiscal,
          row.year2.selfRaised,
          row.year2.outsourcing,
          row.year2.total,
          // 第三年
          row.year3.fiscal,
          row.year3.selfRaised,
          row.year3.outsourcing,
          row.year3.total,
          // 总计
          row.total.fiscal,
          row.total.selfRaised,
          row.total.outsourcing,
          row.total.total
        ];
      });

      // 合并表头和数据
      const excelData = [...headers, ...data];

      // 创建工作表
      const ws = XLSX.utils.aoa_to_sheet(excelData);

      // 设置列宽
      const wscols = [
        { width: 30 }, // 预算科目名称
        { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, // 第一年
        { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, // 第二年
        { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, // 第三年
        { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }  // 总计
      ];
      ws['!cols'] = wscols;

      // 设置单元格合并
      ws['!merges'] = [
        // 合并第一行的表头
        { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } }, // 预算科目名称
        { s: { r: 0, c: 1 }, e: { r: 0, c: 4 } }, // 第一年
        { s: { r: 0, c: 5 }, e: { r: 0, c: 8 } }, // 第二年
        { s: { r: 0, c: 9 }, e: { r: 0, c: 12 } }, // 第三年
        { s: { r: 0, c: 13 }, e: { r: 0, c: 16 } } // 总计
      ];

      // 设置样式
      // 注意：这里的样式设置比较简单，更复杂的样式可能需要其他库支持
      const range = XLSX.utils.decode_range(ws['!ref']);
      for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const cell_address = { c: C, r: R };
          const cell_ref = XLSX.utils.encode_cell(cell_address);

          if (!ws[cell_ref]) continue;

          // 表头单元格样式
          if (R <= 1) {
            if (!ws[cell_ref].s) ws[cell_ref].s = {};
            ws[cell_ref].s.fill = { fgColor: { rgb: "EFEFEF" } };
            ws[cell_ref].s.font = { bold: true };
            ws[cell_ref].s.alignment = { horizontal: "center", vertical: "center" };
            ws[cell_ref].s.border = {
              top: { style: "thin" },
              bottom: { style: "thin" },
              left: { style: "thin" },
              right: { style: "thin" }
            };
          }

          // 将小计和总计行加粗
          const row = this.tableData[R - 2]; // 减去表头的2行
          if (row && (row.isTotal || row.isSubtotal)) {
            if (!ws[cell_ref].s) ws[cell_ref].s = {};
            ws[cell_ref].s.font = { bold: true };
            ws[cell_ref].s.fill = { fgColor: { rgb: "F5F7FA" } };
          }

          // 设置数字列的格式
          if (C > 0) { // 排除第一列（名称列）
            if (!ws[cell_ref].s) ws[cell_ref].s = {};
            ws[cell_ref].s.numFmt = "#,##0.00";
            ws[cell_ref].s.alignment = { horizontal: "right" };
          }

          // 所有单元格添加边框
          if (!ws[cell_ref].s) ws[cell_ref].s = {};
          if (!ws[cell_ref].s.border) {
            ws[cell_ref].s.border = {
              top: { style: "thin" },
              bottom: { style: "thin" },
              left: { style: "thin" },
              right: { style: "thin" }
            };
          }
        }
      }

      // 将工作表添加到工作簿
      XLSX.utils.book_append_sheet(wb, ws, "预算报表");

      // 生成文件名（包含当前日期）
      const now = new Date();
      const dateStr = now.toISOString().split('T')[0]; // YYYY-MM-DD 格式
      const fileName = `预算报表_${dateStr}.xlsx`;

      // 导出文件
      XLSX.writeFile(wb, fileName);
    }
  }
};
</script>

<style scoped>
.edit-button {
  padding: 0;
  margin: 0;
  display: inline-flex;
  justify-content: center;
  align-items: center;
}
.app-container {
  padding: 20px;
}

.clearfix {
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.operations {
  display: flex;
  gap: 10px;
}

.font-bold {
  font-weight: bold;
}

/* 增加表格样式 */
::v-deep .el-table {
  font-size: 14px;
}

::v-deep .el-table th {
  background-color: #f5f7fa !important;
  color: #606266;
  font-weight: bold;
  text-align: center;
}

::v-deep .el-table td {
  padding: 8px 0;
}

/* 设置表格边框 */
::v-deep .el-table--border,
::v-deep .el-table--group {
  border: 1px solid #ebeef5;
}

::v-deep .el-table--border td,
::v-deep .el-table--border th {
  border-right: 1px solid #ebeef5;
}

::v-deep .el-table--border th.is-leaf,
::v-deep .el-table--border td {
  border-bottom: 1px solid #ebeef5;
}
</style>
