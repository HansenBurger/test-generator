import { createRouter, createWebHistory } from 'vue-router'
import Home from '../views/Home.vue'
import Dashboard from '../views/Dashboard.vue'
import OutlineToCases from '../views/OutlineToCases.vue'

const routes = [
    {
        path: '/',
        name: 'Dashboard',
        component: Dashboard
    },
    {
        path: '/outline-generation',
        name: 'OutlineGeneration',
        component: Home
    },
    {
        path: '/outline-to-cases',
        name: 'OutlineToCases',
        component: OutlineToCases
    }
]

const router = createRouter({
    history: createWebHistory(),
    routes
})

export default router



