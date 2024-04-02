import { createRouter, createWebHistory } from 'vue-router'
import principal from './components/caracterizacion/principal';
import editarCaracterizacion from './components/caracterizacion/editarCaracterizacion';
import listaCaracterizados from './components/caracterizacion/listaCaracterizados';
import listaCaracterizadosPorDigitador from './components/caracterizacion/porDigitador';
import dashboard from './components/admin/dashboard';
import miPerfil from './components/miPerfil';
import usuarios from './components/usuarios/usuarios';
import consolidado from './components/caracterizacion/consolidado';
import corregimientos from './components/admin/corregimientos';
import informe from './components/admin/informe';

const routes = [
  {
    path: '/caracterizacion',
    name: 'caracterizacion',
    component: principal
  },
  {
    path: '/dashboard',
    name: 'dashboard',
    component: dashboard
  },
  {
    path: '/:catchAll(.*)',
    redirect: '/'
  },
  {
    path: '/lista-caracterizados',
    name: 'listaCaracterizados',
    component: listaCaracterizados
  },
  {
    path: '/editar-caracterizacion/:identificacion_editar',
    name: 'editarCaracterizacion',
    component: editarCaracterizacion,
    props: true
  },
  {
    path: '/mi-perfil',
    name: 'miPerfil',
    component: miPerfil
  },
  {
    path: '/usuarios',
    name: 'usuarios',
    component: usuarios
  },
  {
    path: '/consolidado',
    name: 'consolidado',
    component: consolidado
  },
  {
    path: '/corregimientos',
    name: 'corregimientos',
    component: corregimientos
  },
  {
    path: '/informe',
    name: 'informe',
    component: informe
  },
  {
    path: '/lista-caracterizados-digitador',
    name: 'listaCaracterizadosPorDigitador',
    component: listaCaracterizadosPorDigitador
  },
]

const router = createRouter({
  history: createWebHistory(process.env.BASE_URL),
  routes
})

export default router