<template>
  <q-layout view="hHh Lpr lFf">
    <q-page-container class="bg-grey-5">
      <q-page padding class="row items-center justify-center">
        <div class="row full-width">
          <div class="col-md-4 offset-md-4 col-xs-12 q-pl-md q-pr-md q-pt-sm">
            <q-card flat class="bg-white text-black">
              <div class="row">
               

                <div class="col-md-12 col-xs-12">
                  <div class="q-pa-md">
                    <div
                      class="text-h6 q-pb-md text-blue-8 text-center text-weight-bolder"
                    >
                      Login WCS SYSYEM(MU)
                    </div>
                    <q-form
                      @submit="onSubmit"
                      @reset="onReset"
                      class="q-gutter-md"
                    >
                      <q-input
                        filled
                        v-model="name"
                        label="Username"
                        lazy-rules
                        :rules="[
                          (val) =>
                            (val && val.length > 0) || 'Please type Username',
                        ]"
                      />

                      <q-input
                        filled
                        type="password"
                        v-model="password"
                        label="Password"
                        
                
                      />
                      
<br>
                       <q-select filled v-model="org" :options="options" label="Chose Org" />

                      <div>
                        <q-btn  label="Login" type="submit" color="primary" />
                       
                        <q-btn
                          label="Reset"
                          type="reset"
                          color="primary"
                          flat
                          class="q-ml-sm"
                          
                        />
                          <div
                      class=" q-pb-md text-grey-6 text-right "
                    >
                      V3.2.1
                    </div>

                        
                      </div>
                    </q-form>
                  </div>
                </div>
              </div>
            </q-card>
          </div>
        </div>
      </q-page>
    </q-page-container>
  </q-layout>
</template>

<script>

import axios from 'axios'
import { useQuasar } from 'quasar'
import {ref} from 'vue'
import { onBeforeUnmount } from 'vue'
import {
    Cookies
} from 'quasar'
export default {
  data() {
    return {
      name: ref(null),
      password: ref(null),
      options:["G1","G2","G3","G4"],
      org:ref("")
    };
  },

  methods: {
    onSubmit() {
      if(this.org == null || this.org == "" || this.org == undefined){
            this.showNotif3()
      }else{
      const params = new FormData()
      params.append('name',this.name)
      params.append('password',this.password)
       for (var pair of params.entries()) {
          console.log(pair[0] + ", " + pair[1]);
        }
      axios({
                method: 'post',
                url: this.$api_url + '/mylogin.php/test',
                data: params
            }).then((resp)=>{
              console.log(resp.data)
               console.log(resp.data.data[0].CHG_TARGET)
                 if(resp.data.status == true){
                    if(resp.data.status_login == true){
                     
                        console.log('success')
                        console.log(resp.data.status_login)
                        this.$q.localStorage.set('username', resp.data.data[0].USER_ID);
                        this.$q.localStorage.set('Password', resp.data.data[0].USER_PASSWORD);
                        this.$q.localStorage.set('org', this.org);
                        this.$q.localStorage.set('status', resp.data.data[0].CHG_TARGET);
                        this.$q.localStorage.set('login_status', resp.data.status_login);
                        const option = {
                            secure: false,
                            expires: '12h' // in 15 minutes, 10 seconds
                        }
                       
                        Cookies.set('name', this.name, option)
                        Cookies.set('password', this.password, option)
                        
                       
                  
                         this.$router.push({
                            path: '/first/main'
                        }) 
                      
                    }else{
                        console.log('username or password incorrect')
                         this.showNotif2()
                    }
                }else{
                    console.log('username or password incorrect')
                     this.showNotif2()
                     this.showLoading();
                }
               
              
            }).catch((error)=>{
                console.log(error)
            })
      }
    },
    onReset() {
      this.name = null;
      this.password = null;
    },
  },
  setup () {
    const $q = useQuasar()
    let timer

    onBeforeUnmount(() => {
      if (timer !== void 0) {
        clearTimeout(timer)
        $q.loading.hide()
      }
    })

    return {
      showNotif () {
        $q.notify({
          message: 'Success',
          position:"center",
          icon: 'announcement'
        })
      },
      showNotif2 () {
        $q.notify({
          message: 'Username or password incorrect',
          position:"center",
          icon: 'announcement',
          color:"red-9"
        })
      },
       showNotif3 () {
        $q.notify({
          message: 'Please Chose org',
          position:"center",
          icon: 'announcement',
          color:"red-9"
        })
      },
       showLoading () {
        $q.loading.show()

        // hiding in 2s
        timer = setTimeout(() => {
          $q.loading.hide()
          timer = void 0
        }, 2000)
      }

    }
  }
};
</script>