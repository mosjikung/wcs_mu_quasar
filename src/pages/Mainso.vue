<template>

 <div class="q-pa-md" style="max-width: 100%">
  
    <div class="q-gutter-md">
  <q-card class="my-card" style="max-width:100%">
    <div class="row">
        <div class="col-4">

        </div>
            <div class="col-2">
                <q-input v-model="SO_NO_DOC" filled type="search" dense  @keydown.enter="checksearch()" label="So">
                    <template v-slot:append>
                    <q-icon name="search" />
                    </template>
                </q-input>
            </div>
    
       
        <div class="col-1">
    

        </div>
        <div class="col-2">
            
        <q-btn icon="search"  color="cyan-8" size="md" label="search" @click="checksearch()"></q-btn>
        
        </div>

        
        </div><br><br>
         <div class="row">
        <div class="col-4">

        </div>
            <div class="col-2">
                   <q-select 
        filled
        v-model="multiple"
        dense
        :options="options"
        label="Chose Report"
      />
            </div>
    
       
        <div class="col-1">
    

        </div>
        <div class="col-2">
            
        <q-btn icon="search"  color="primary" size="md" label="Export" @click="Check_value_choice"></q-btn>
        </div>

        
        </div>
      

        </q-card>
        </div><br>
      
        <div class="row">
        <div class="col-4">
          <q-btn
      size="md"
      class="q-px-xl q-py-xs "
      color="primary"
      label="new"
      @click="$router.push('/first/new')"
    />
   
       </div>
        </div>
         <br>
        <div class="q-gutter-md">
          <div class="row">
                <div class="col-12">
                <div class="q-pt-md" style="margin-top: 0px;padding-top: 0px;">
                                        <q-table :rows="rowshow" :columns="columnsshow" :rows-per-page-options="[10]" dense row-key="MU_SEQ" style="min-height: 340px;">
                                            <template v-slot:body="props">
                                                <q-tr :props="props">
                                                  
                                                        <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                                                        <!--   -->
                                                      
                                               
                                                      <q-td key="MU_SEQ" :props="props">
                                                      <q-input input-class="text-center"  v-model="props.row.MU_SEQ"  dense disable />

                                                    </q-td>
                                                     <q-td key="GM_NO" :props="props">
                                                      <q-input input-class="text-center"  v-model="props.row.GM_NO"  dense disable />

                                                    </q-td>
                                                     <q-td key="no" :props="props">
                                                      <q-input input-class="text-center" v-model="props.row.SO_NO_DOC"  dense disable />

                                                    </q-td>
                                                     <q-td key="MU_DATE" :props="props">
                                                      <q-input input-class="text-center" v-model="props.row.MU_DATE"  dense disable />

                                                    </q-td>
                                                    
                                                    
                                                  
                                                     <q-td key="year" :props="props">
                                                     <q-input input-class="text-center" v-model="props.row.SO_YEAR"  dense disable />
                                                    </q-td>
                                                  
                                                     <q-td key="style" :props="props">
                                                        <q-input input-class="text-center" v-model="props.row.STYLE" dense disable />
                                                    </q-td>
                                                    <q-td key="customer" :props="props"> <!--key = name.columne -->
                                                       <q-input input-class="text-center"    v-model="props.row.CUST"  dense disable />
                                                    </q-td>
                                                    <q-td key="customer" dense :props="props"> <!--key = name.columne -->
                                                         <Detail :detail="props.row.SO_NO" :detail3="props.row.SO_YEAR" :detail4="props.row.SO_NO_DOC"
                                                         :detail5="props.row.CPART" :detail6="props.row.MU_SEQ" @show="getshow" @clicked="submit"></Detail>
                                                    </q-td>
                                                   
                                                  

                                                    
                                          
                                                </q-tr>
                                            </template>
                                        </q-table>
                                    </div>
                </div>

                

           </div> <!-- row -->
    
    </div>
    </div>

     <q-dialog
      v-model="open_all"
      transition-show="rotate"
      transition-hide="rotate"
    >
      <q-card style="min-width: 80%">
        <q-card-section>
          <div class="text-h6">Export Excel Spread Loss Data</div>
          <q-card-actions align="center">
          <ExcelALL></ExcelALL>
          </q-card-actions>
        </q-card-section>

        <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
          
        </div>
      </q-card>
    </q-dialog>
    <q-dialog
      v-model="open_daily"
      transition-show="rotate"
      transition-hide="rotate"
    >
      <q-card style="min-width:80%">
        <q-card-section>
          <div class="text-h6">Export Excel Daily</div>
          <q-card-actions align="center">
          <ExcelDaily></ExcelDaily>
          </q-card-actions>
        </q-card-section>

        <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
          
        </div>
      </q-card>
    </q-dialog>
      <q-dialog
      v-model="open_weekly"
      transition-show="rotate"
      transition-hide="rotate"
    >
      <q-card style="min-width:100%">
        <q-card-section>
          <div class="text-h6">Export Excel Weekly</div>
          <q-card-actions align="center">
          <ExcelWeekly></ExcelWeekly>
          </q-card-actions>
        </q-card-section>

        <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
          
        </div>
      </q-card>
    </q-dialog>
     <q-dialog
      v-model="open_weekly_wla"
      transition-show="rotate"
      transition-hide="rotate"
    >
      <q-card style="min-width:100%">
        <q-card-section>
          <div class="text-h6">Export Excel Weekly MU WLA-1</div>
          <q-card-actions align="center">
         <Excelwla></Excelwla>
          </q-card-actions>
        </q-card-section>

        <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
          
        </div>
      </q-card>
    </q-dialog>
       <q-dialog
      v-model="open_weekly_data_wla"
      transition-show="rotate"
      transition-hide="rotate"
    >
      <q-card style="min-width:100%">
        
         <ExcelDataWLa></ExcelDataWLa>
       
      </q-card>
    </q-dialog>

  
       <q-dialog
      v-model="open_daily_data_31"
      transition-show="rotate"
      transition-hide="rotate"
    >
      <q-card style="min-width:100%">
       
         <ExcelSla31></ExcelSla31>
      
      </q-card>
    </q-dialog>

     <q-dialog
      v-model="open_daily_data_32"
      transition-show="rotate"
      transition-hide="rotate"
    >
      <q-card style="min-width:100%;">
      
      
         <ExcelSla32></ExcelSla32>
       
      </q-card>
    </q-dialog>

    <q-dialog
      v-model="open_analysis"
      transition-show="rotate"
      transition-hide="rotate"
    >
      <q-card style="min-width:100%">
        <q-card-section>
          <div class="text-h6">Weekly Analysis</div>
          <q-card-actions align="center">
         <ExcelWeeklyAnaly></ExcelWeeklyAnaly>
          </q-card-actions>
        </q-card-section>

        <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
          
        </div>
      </q-card>
    </q-dialog>
    
    <q-dialog
      v-model="open_setting"
      transition-show="rotate"
      transition-hide="rotate"
    >
      <q-card style="min-width:70%; magrin: 0px; padding: 0px; padding-bottom: 30px;">
          <Setting></Setting>
       
      </q-card>
    </q-dialog>

    <q-dialog
      v-model="open_setting_history"
      transition-show="rotate"
      transition-hide="rotate"
    >
      <q-card style="min-width:100%; magrin: 0px; padding: 0px; padding-bottom: 30px;">
          <Settinghis></Settinghis>
       
      </q-card>
    </q-dialog>

    <q-dialog
      v-model="open_test"
      transition-show="rotate"
      transition-hide="rotate"
    >
      <q-card style="min-width:100%; magrin: 0px; padding: 0px; padding-bottom: 30px;">
          <Test></Test>
       
      </q-card>
    </q-dialog>
</template>

<script>
import {ref} from 'vue'
import {date} from 'quasar'
import axios from 'axios'
import Detail from '../components/Detail.vue'
import ExcelDaily from '../components/Excel_Daily.vue'
import ExcelALL from '../components/ExcelAll.vue'
import ExcelWeekly from '../components/ExcelWeekly.vue'
import Excelwla from '../components/Excelwla.vue'
import ExcelDataWLa from '../components/ExcelDataWLa.vue'
import ExcelSla31 from '../components/ExcelSLA31.vue'
import ExcelSla32 from '../components/ExcelSLA32.vue'
import ExcelWeeklyAnaly from '../components/ExcelWeeklyAnaly.vue'
import Setting from '../components/Setting.vue'
import Settinghis from '../components/Setting_history.vue'
import Test from '../components/Test_components.vue'
import { useQuasar, QSpinnerGears } from 'quasar'
import { onBeforeUnmount } from 'vue'
import * as Excel from "exceljs";
import { saveAs } from "file-saver";
export default {
    
    setup(){
       const $q = useQuasar()
    let timer

    onBeforeUnmount(() => {
      if (timer !== void 0) {
        clearTimeout(timer)
        this.$q.loading.hide()
      }
    })
    
        return{
            date: ref('2021-10-21')
        }
    },
    mounted() {
      let login_status = this.$q.localStorage.getItem('login_status')
      console.log(login_status)
     
      let org = this.$q.localStorage.getItem('org')
      this.org = org
      if(login_status !== true || login_status  == null){
        this.$q.notify({
          message: "Please Login",
          position:"center",
          color: "red-9",
        });
       
        this.$router.push('/login')
      }
    },
     
        data(){
            return{
                search:"",
                 multiple: ref(null),
                 options:['Spread Loss Data','Daily Report','Weekly Report MU-SLA2','Weekly Report WLA'
                 ,'Weekly Report WLA Data','Weekly Mu Sla3-1','Weekly Mu Sla3-2','Weekly Analysis','Setting'
                 ,'Setting History Performance','Test'],

                
                columnsshow:[
                    {
                        name:"MU_SEQ",
                        label:"MU_SEQ",
                        align:"center",
                        field:"MU_SEQ"
                    },
                     {
                        name:"GM_NO",
                        label:"GM NO",
                        align:"center",
                        field:"GM NO"
                    },
                    {
                        name:"no",
                        label:"SO NO DOC",
                        align:"center",
                        field:"SO_NO_DOC"
                    },
                     {
                        name:"MU_DATE",
                        label:"MU DATE",
                        align:"center",
                        field:"MU_DATE"
                    },
                    {
                        name:"year",
                        label:"Year",
                        align:"center",
                        field:"SO_YEAR"
                    },
                    {
                        name:"style",
                        label:"Style",
                        align:"center",
                        field:"STYLE"
                    },
                    {
                        name:"customer",
                        label:"CUSTOMER",
                        align:"center",
                        field:"CUST"
                    },
                     {
                        name:"detail",
                        label:"Detail",
                        align:"center",
                        field:"detail"
                    },

                    

                ],
                rowshow:[
                   
                ],
                 rowexport: [],
                 start:ref(""),
      end:ref(""),
      org:ref(""),
      option_org:["G1","G2","G3","G4"],
      status_export:1,
      status_export_bottom:0,
      post_excel_before:[],
      post_excel:[],
      post_excel_after:[],
      open_all:false,
      open_daily:false,
      open_weekly:false,
      open_weekly_wla:false,
      open_weekly_data_wla:false,
      open_daily_data_31:false,
      open_daily_data_32:false,
      open_setting:false,
      open_analysis:false,
      open_setting_history:false,
      open_test:false,

            }
            
        },
        components:{
           Detail,
           ExcelALL,
           ExcelDaily,
           ExcelWeekly,
           Excelwla,
           ExcelDataWLa,
           ExcelSla31,
           ExcelSla32,
           ExcelWeeklyAnaly,
           Setting,
           Settinghis,
           Test,
        },
        
        methods: {
            checksearch(){
                 this.$q.loading.show({
          spinner: QSpinnerGears,
          spinnerColor: 'wthite',
          spinnerSize: 180,
          backgroundColor: 'black',
         
        })

       
                const params = new FormData()
                params.append('SO_NO_DOC',this.SO_NO_DOC)
                this.rowshow=[];

                 axios({
                    method:'post',
                    url:this.$api_url + '/mainso.php/checksearch',
                    data:params
                }).then((resp)=>{
                    console.log(resp.data)
                     if(resp.data.data.length >0){
                     resp.data.data.forEach(e =>{
                         this.rowshow.push({
                          MU_SEQ : e.MU_SEQ,
                          SO_NO : e.SO_NO,
                          SO_NO_DOC: e.SO_NO_DOC,
                          SO_YEAR: e.SO_YEAR,
                          MU_DATE: e.MU_DATE,
                          GM_NO: e.GM_NO,
                          STYLE: e.STYLE_REF,
                          CUST: e.CUSTOMER_NAME,
                       


                         })
                     })
                   
                  }else{
                     this.$q.notify({
                      message: "Wrong SO",
                      position:"center",
                      color: "red-9",
        });
                  }
                
                }) 
                 // hiding in 3s
        this.timer = setTimeout(() => {
          this.$q.loading.hide()
         this.timer = void 0
        }, 800)
            },
            goto(val){
                console.log(val)

            },
            getshow(value){
                console.log(this.rowshow)
                this.rowshow=[]
            
             
              this.rowshow = value
            
            },
            Check_value_choice(){
            if(this.multiple == "Spread Loss Data"){
              this.open_all = true
            }else if(this.multiple == "Daily Report"){
              this.open_daily = true
            }else if(this.multiple == "Weekly Report MU-SLA2"){
              this.open_weekly = true
            }else if(this.multiple == "Weekly Report WLA"){
              this.open_weekly_wla = true
            }else if(this.multiple == "Weekly Report WLA Data"){
              this.open_weekly_data_wla = true
            }else if(this.multiple == "Weekly Mu Sla3-1"){
              this.open_daily_data_31 = true
            }else if(this.multiple == "Weekly Mu Sla3-2"){
              this.open_daily_data_32 = true
            }else if(this.multiple == "Setting"){
              this.open_setting = true
            }else if(this.multiple == "Weekly Analysis"){
              this.open_analysis = true
            }else if(this.multiple == "Setting History Performance"){
              this.open_setting_history = true
            }else if(this.multiple == "Test"){
              this.open_test= true
            }

            },
            
            Check_choice(){
            
               for(var ax = 0; ax<this.multiple.length; ax++){
                if(this.multiple[ax] == "ALL"){
                  this.Export();
                  
                }else{
                  this.status_export = 2
                  
                }

                
                if(this.multiple[ax] == "Woven"){
                  this.chosen = 1
                }
                if(this.multiple[ax] == "Light-Middle Weight"){
                  this.chosen = 2
                }
                if(this.multiple[ax] == "Fleece"){
                  this.chosen = 3
                }
                if(this.multiple[ax] == "Special"){
                  this.chosen = 4
                }
                console.log(this.start_date)
                console.log(this.end_date)
                console.log(this.chosen)
                 this.post_excel.push({
                start_date:this.start_date,
                end_date:this.end_date,
                f_type:this.chosen,
                }) 
              } 
            }
            ,
             
     
       
       
       },

      watch:{
        start(val){
          this.start_date =""
          let day = val.substring(0, 2);
          let month = val.substring(3, 5);
          let year = val.substring(6, 10);
          this.start_date = year +"/"+ month +"/"+ day
          
        },
        end(val){
          this.end_date = ""
           let day = val.substring(0, 2);
          let month = val.substring(3, 5);
          let year = val.substring(6, 10);
          this.end_date = year +"/"+ month +"/"+ day
          
        }
      }
}
</script>

<style lang="sass">
.my-card
  padding-top: 50px
  padding-right: 30px
  padding-bottom: 50px
  padding-left: 80px
</style>