import Layout from '@/layout'
// 预算管理相关路由
export const budgetRoutes = {
  path: '/budget',
  component: Layout,
  redirect: '/budget/report',  // 修改为具体的子路由路径
  name: 'Budget',
  meta: {
    title: '预算管理',
    icon: 'money'
  },
  children: [
    {
      path: 'report',
      component: () => import('@/views/budget/report/index'),
      name: 'BudgetReport',
      meta: { title: '预算报表', icon: 'table' }
    },
    // 可以添加更多预算相关的子路由...
  ]
};
