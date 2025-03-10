import request from '@/utils/request';

// 获取预算报表数据
export function getBudgetReport(projectId) {
  return request({
    url: `/budget/report/${projectId}`,
    method: 'get'
  });
}

// 保存预算报表数据
export function saveBudgetReport(data) {
  return request({
    url: '/budget/report/save',
    method: 'post',
    data: data
  });
}

// 导出预算报表为Excel
export function exportBudgetReport(projectId) {
  return request({
    url: `/budget/report/export/${projectId}`,
    method: 'get',
    responseType: 'blob'
  });
}
