<template>
  <q-layout view="lHh Lpr lFf">
     <q-header elevated class="bg-deep-purple-10">
      <q-toolbar>
        <q-btn
          flat
          dense
          round
          icon="menu"
          aria-label="Menu"
          @click="toggleLeftDrawer"
        />

        <q-toolbar-title>
          WCS SYSTEM (MU)
        </q-toolbar-title>
            <q-btn
               class="q-ma-xs"
              icon="arrow_back"
              type="submit"
              size="md"
              label=""
              stack
              
              
              @click="Back()"
            >
               <q-tooltip
          transition-show="flip-right"
          transition-hide="flip-left"
        >
          Back to main Page
        </q-tooltip>
            </q-btn>
        <q-btn icon="logout" size="md" @click="log_out()">

        </q-btn>
         
        <div></div>
        <div class="q-pa-md q-gutter-sm">
        <q-dialog
          v-model="basic"
          transition-show="rotate"
          transition-hide="rotate"
        >
          <q-card style="min-width: 30%">
            <q-card-section>
              <div class="text-h6">Confirm Logout</div>
              <q-card-actions align="right">
                <q-btn color="green-4" size="lg" icon="done" @click="confirm_log_out()"/>
                <q-btn color="negative" size="lg" icon="cancel"  @click="cancel_log_out()" />
              </q-card-actions>
            </q-card-section>

           
          </q-card>
        </q-dialog>
      </div>
      </q-toolbar>
    </q-header>

    

    <q-page-container>
      <router-view />
    </q-page-container>
  </q-layout>
</template>

<script>
import EssentialLink from 'components/EssentialLink.vue'
import axios from 'axios'
import { useQuasar } from 'quasar'

import { onBeforeUnmount } from 'vue'
import {
    Cookies
} from 'quasar'
const linksList = [
  {
    title: 'Docs',
    caption: 'quasar.dev',
    icon: 'school',
    link: 'https://quasar.dev'
  },
  {
    title: 'Github',
    caption: 'github.com/quasarframework',
    icon: 'code',
    link: 'https://github.com/quasarframework'
  },
  {
    title: 'Discord Chat Channel',
    caption: 'chat.quasar.dev',
    icon: 'chat',
    link: 'https://chat.quasar.dev'
  },
  {
    title: 'Forum',
    caption: 'forum.quasar.dev',
    icon: 'record_voice_over',
    link: 'https://forum.quasar.dev'
  },
  {
    title: 'Twitter',
    caption: '@quasarframework',
    icon: 'rss_feed',
    link: 'https://twitter.quasar.dev'
  },
  {
    title: 'Facebook',
    caption: '@QuasarFramework',
    icon: 'public',
    link: 'https://facebook.quasar.dev'
  },
  {
    title: 'Quasar Awesome',
    caption: 'Community Quasar projects',
    icon: 'favorite',
    link: 'https://awesome.quasar.dev'
  }
];

import { defineComponent, ref } from 'vue'

export default defineComponent({
  name: 'MainLayout',

  components: {
    
  },

  setup () {
    const leftDrawerOpen = ref(false)
    
    return {
      essentialLinks: linksList,
      leftDrawerOpen,
      toggleLeftDrawer () {
        leftDrawerOpen.value = !leftDrawerOpen.value
      }
    }
  },
  data(){
    return{
      basic:false
    }
  },
  methods: {
    log_out(){
      this.basic=true
    },
    confirm_log_out(){
      this.$router.push('/login')
      this.$q.localStorage.clear()
    },
    cancel_log_out(){
      this.basic=false
    },
    Back(){
     this.$router.push('/first/main');
    },
  },
})
</script>
