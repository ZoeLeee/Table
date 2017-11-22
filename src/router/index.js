import Vue from 'vue'
import Router from 'vue-router'
import InfoTable from '@/components/info-table';

Vue.use(Router)

export default new Router({
  routes: [
    {
      path:"/",
      component:InfoTable
    }
  ]
})
