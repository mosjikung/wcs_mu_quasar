
const routes = [
  {
    path: '/first',
    component: () => import('layouts/MainLayout.vue'),
    children: [
      { path: '', component: () => import('pages/Index.vue') },
      { path: '/first/new', component: () => import('pages/Somain.vue') },
      { path: '/first/main', component: () => import('pages/Mainso.vue') },
      { path: '/first/test', component: () => import('pages/Test.vue') },
      { path: '/first/dialog', component: () => import('pages/Dialog.vue') },
      { path: '/first/test2', component: () => import('pages/Test2.vue') },
      { path: '/first/table', component: () => import('pages/Table.vue') },
      { path: '/first/new2', component: () => import('pages/New.vue') },
      
    ]
  },
  {
    path: '/',
    component: () => import('pages/Login.vue'),
    children: [
      { path: '', component: () => import('pages/Login.vue') },
      { path: '/login', component: () => import('pages/Login.vue') },
    ]
  },

  // Always leave this as last one,
  // but you can also remove it
  {
    path: '/:catchAll(.*)*',
    component: () => import('pages/Error404.vue')
  }
]

export default routes
