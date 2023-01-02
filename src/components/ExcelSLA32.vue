<template>
<div class="bg-blue-10 text-white">
      <q-toolbar>
        <q-btn flat round dense />

        <q-toolbar-title>Weekly Mu SLA 3-2</q-toolbar-title>

        <q-btn icon="cancel" flat fixed-right color="white" v-close-popup />
      </q-toolbar>
    </div>
    <br>
  <div class="row justify-center">
      <q-input class="q-pa-xl" filled v-model="start">
      <template v-slot:append>
        <q-icon name="event" class="cursor-pointer">
          <q-popup-proxy
            ref="qDateProxy"
            transition-show="scale"
            transition-hide="scale"
          >
            <q-date v-model="start" mask="DD/MM/YYYY">
              <div class="row items-center justify-end">
                <q-btn v-close-popup label="Close" color="primary" flat />
              </div>
            </q-date>
          </q-popup-proxy>
        </q-icon>
      </template>
    </q-input>
   
    <q-input class="q-pa-xl" filled v-model="end">
      <template v-slot:append>
        <q-icon name="event" class="cursor-pointer">
          <q-popup-proxy
            ref="qDateProxy"
            transition-show="scale"
            transition-hide="scale"
          >
            <q-date v-model="end" mask="DD/MM/YYYY">
              <div class="row items-center justify-end">
                <q-btn v-close-popup label="Close" color="primary" flat />
              </div>
            </q-date>
          </q-popup-proxy>
        </q-icon>
      </template>
    </q-input>
  
  <div class="myclass">  <q-btn
      size="lg"
      dense
      color="primary"
      label="Export"
      @click="exportexcel()"
    /></div>
  
    <br>
    <br>
 
<div class="q-pa-xl" style="min-width:100%;" >
  <q-banner inline-actions rounded class="bg-orange-4 text-white" >
    <h5 class="q-mt-xs" >Woven</h5>
      <template v-slot:action>
        
      </template>
    </q-banner> 
     <q-card class="my-card" >
       
      <q-card-section>
       <div class ="col-12">
  
  </div>
    <div class="col-12" style="min-width:100%; min-height:600px;" id="chartDom"></div>
      </q-card-section>
    </q-card>
</div>

<div class="q-pa-xl" style="min-width:100%;" >
  <q-banner inline-actions rounded class="bg-light-blue-8  text-white">
    <h5 class="q-mt-xs" >Light & Middle Weight</h5>
      <template v-slot:action>
        
      </template>
    </q-banner> 
     <q-card class="my-card" >
       
      <q-card-section>
       <div class ="col-12">
  
  </div>
    <div class="col-12" style="min-width:100%; min-height:600px;" id="chartDom2"></div>
      </q-card-section>
    </q-card>
</div>

<div class="q-pa-xl" style="min-width:100%;" >
  <q-banner inline-actions rounded class="bg-pink-2 text-white" >
    <h5 class="q-mt-xs" >Fleece</h5>
      <template v-slot:action>
        
      </template>
    </q-banner> 
     <q-card class="my-card" >
       
      <q-card-section>
       <div class ="col-12">
  
  </div>
    <div class="col-12" style="min-width:100%; min-height:600px;" id="chartDom3"></div>
      </q-card-section>
    </q-card>
</div>


<div class="q-pa-xl" style="min-width:100%;" >
  <q-banner inline-actions rounded class="bg-grey-4 text-white" >
    <h5 class="q-mt-xs" >Special</h5>
      <template v-slot:action>
        
      </template>
    </q-banner> 
     <q-card class="my-card" >
       
      <q-card-section>
       <div class ="col-12">
  
  </div>
    <div class="col-12" style="min-width:100%; min-height:600px;" id="chartDom4"></div>
      </q-card-section>
    </q-card>
</div>


    
  </div>
  

</template>


<script>
import { ref } from "vue";
import { date } from "quasar";
import axios from "axios";
import { useQuasar, QSpinnerGears } from "quasar";
import * as Excel from "exceljs";
import { saveAs } from "file-saver";
import moment from "moment";
import * as echarts from 'echarts';
export default {
  data(){
    return{
      start:"",
      end:"",
      first_day:"",
      last_day:"",
      open_data:false,
      org:[],
      
      rowexport:[],
      allrowexport:[],
      Woven_endloss:0.15,
      Woven_spliceloss:0.20,
      Woven_cutoutloss:0.00,
      Woven_remnantsloss:0.08,
      Woven_total:"",
      Light_endloss:0.15,
      Light_spliceloss:0.20,
      Light_cutoutloss:0.00,
      Light_remnantsloss:0.14,
      Light_total:"",
      Fleece_endloss:0.20,
      Fleece_spliceloss:0.20,
      Fleece_cutoutloss:0.00,
      Fleece_remnantsloss:0.25,
      Fleece_total:"",
      Special_endloss:0.23,
      Special_spliceloss:0.20,
      Special_cutoutloss:0.00,
      Special_remnantsloss:0.24,
      Special_total:"",
      date_use_data:[],
      all_data_arr:[],
      graph_date:{},
      graph_data:{},
      date_data_for_graph1:[],
      date_data_for_graph2:[],
      date_data_for_graph3:[],
      date_data_for_graph4:[],



      data_graph1:[],
      data_graph2:[],
      data_graph3:[],
      data_graph4:[],

      data_target_graph1:[],
      data_target_graph2:[],
      data_target_graph3:[],
      data_target_graph4:[],


      
      ttl_marker_endless_woven:[],
      ttl_marker_endless_light:[],
      ttl_marker_endless_fleece:[],
      ttl_marker_endless_special:[],
      
      ttl_endloss_yd_woven:[],
      ttl_endloss_yd_light:[],
      ttl_endloss_yd_fleece:[],
      ttl_endloss_yd_special:[],

      ttl_splice_woven:[],
      ttl_splice_light:[],
      ttl_splice_fleece:[],
      ttl_splice_special:[],


      ttl_cut_woven:[],
      ttl_cut_light:[],
      ttl_cut_fleece:[],
      ttl_cut_special:[],


      ttl_rem_woven:[],
      ttl_rem_light:[],
      ttl_rem_fleece:[],
      ttl_rem_special:[],


      ttl_woven:[],
      ttl_light:[],
      ttl_fleece:[],
      ttl_special:[],

      end_target_woven:[],
      splice_target_woven:[],
      cut_target_woven:[],
      rem_target_woven:[],

      end_target_light:[],
      splice_target_light:[],
      cut_target_light:[],
      rem_target_light:[],

      end_target_fleece:[],
      splice_target_fleece:[],
      cut_target_fleece:[],
      rem_target_fleece:[],

      end_target_special:[],
      splice_target_special:[],
      cut_target_special:[],
      rem_target_special:[],

      total_ac_end_arr:[],
      total_ac_splice_arr:[],
      total_ac_cut_arr:[],
      total_ac_rem_arr:[],
      total_ac_all_arr:[],

      total_audit:[],
      total_target:[],
      total_end_target:[],
      total_splice_target:[],
      total_cut_target:[],
      total_rem_target:[],

      sum_all_4_ttl_marker:[],

      ac_end_loss:[],
      ac_splice_loss:[],
      ac_cut_out_loss:[],
      ac_remnants_loss:[],

      ac_end_sum:[],
      ac_splice_sum:[],
      ac_cut_sum:[],
      ac_rem_sum:[],

      all_find_max:[],

      arr_end_loss_1_4:[],

      last_total_spread_end:[],
      last_total_spread_splice:[],
      last_total_spread_cut:[],
      last_total_spread_rem:[],

      bx_summary_ttl_marker_woven:[],
      bx_summary_ttl_marker_light:[],
      bx_summary_ttl_marker_fleece:[],
      bx_summary_ttl_marker_special:[],

      end_loss_woven:"",
        end_loss_light:"",
        end_loss_fleece:"",
        end_loss_special:"",

        splice_loss_woven:"",
        splice_loss_light:"",
        splice_loss_fleece:"",
        splice_loss_special:"",

        cut_loss_woven:"",
        cut_loss_light:"",
        cut_loss_fleece:"",
        cut_loss_special:"",

        rem_loss_woven:"",
        rem_loss_light:"",
        rem_loss_fleece:"",
        rem_loss_special:"",

        total_woven:"",
        total_light:"",
        total_fleece:"",
        total_special:"",

        total_spread_loss_woven_data:[],
        total_spread_loss_woven_data_find_0:[],
        total_spread_loss_woven_target:[],
        total_spread_loss_light_data:[],
        total_spread_loss_light_target:[],
        total_spread_loss_fleece_data:[],
        total_spread_loss_fleece_target:[],
        total_spread_loss_special_data:[],
        total_spread_loss_special_target:[],

        sum_widthloss:[],
        sum_ttl_marker:[],
        sum_width_loss_ttl_marker : [],

        target_width:[],
        acc_width_arr:[],
        end_loss_target_arr:[],
        count_result_arr:[],

        result_per_week:[],
 	      result_max_woven:[],
        result_max_light:[],
        result_max_fleece:[],
        result_max_special:[],

        keep_history_data:[],

        year:[],
        end_loss_value:[],
        splice_loss_value:[],
        cut_out_loss_value:[],
        remnants_loss_value:[],
        total_spread_loss_per_value:[],
        total_width_loss_per_value:[],

       

        

      

      end_loss:[
        {
        p:0.23
        },
         {
        p:0.32
        },
         {
        p:0.35
        },
         {
        p:0.45
        },

        ],

         splice_loss:[
        {
        p:0.10
        },
         {
        p:0.10
        },
         {
        p:0.10
        },
         {
        p:0.10
        },

        ],
          cut_out_loss:[
        {
        p:0.00
        },
         {
        p:0.00
        },
         {
        p:0.00
        },
         {
        p:0.00
        },

        ],

          remnants_loss:[
        {
        p:0.05
        },
         {
        p:0.08
        },
         {
        p:0.11
        },
         {
        p:0.10
        },

        ],

      arr_left_colume:[
         {
          name_col:"A"
        },
         {
          name_col:"B"
        },
         {
          name_col:"C"
        },
         {
          name_col:"D"
        },
         {
          name_col:"E"
        },
         {
          name_col:"F"
        },
         {
          name_col:"G"
        },
         {
          name_col:"H"
        },
         {
          name_col:"I"
        },
         {
          name_col:"J"
        },
         {
          name_col:"K"
        },
         {
          name_col:"L"
        },
         {
          name_col:"M"
        },
         {
          name_col:"N"
        },
         {
          name_col:"O"
        },
         {
          name_col:"P"
        },
         {
          name_col:"Q"
        },
         {
          name_col:"R"
        },
         {
          name_col:"S"
        },
       
      
      ],

      arr_colume:[
        {
          name_col:"X"
        },
         {
          name_col:"Y"
        },
         {
          name_col:"Z"
        },
         {
          name_col:"AA"
        },
         {
          name_col:"AB"
        },
         {
          name_col:"AC"
        },
         {
          name_col:"AD"
        },
         {
          name_col:"AE"
        },
         {
          name_col:"AF"
        },
         {
          name_col:"AG"
        },
         {
          name_col:"AH"
        },
         {
          name_col:"AI"
        },
         {
          name_col:"AJ"
        },
         {
          name_col:"AK"
        },
         {
          name_col:"AL"
        },
         {
          name_col:"AM"
        },
         {
          name_col:"AN"
        },
         {
          name_col:"AO"
        },
         {
          name_col:"AP"
        },
         {
          name_col:"AQ"
        },
         {
          name_col:"AR"
        },
         {
          name_col:"AS"
        },
         {
          name_col:"AT"
        },
         {
          name_col:"AU"
        },
         {
          name_col:"AV"
        },
         {
          name_col:"AW"
        },
         {
          name_col:"AX"
        },
         {
          name_col:"AY"
        },
         {
          name_col:"AZ"
        },
         {
          name_col:"BA"
        },
         {
          name_col:"BB"
        },
         {
          name_col:"BC"
        },
         {
          name_col:"BD"
        },
         {
          name_col:"BE"
        },
         {
          name_col:"BF"
        },
         {
          name_col:"BG"
        },
         {
          name_col:"BH"
        },
         {
          name_col:"BI"
        },
         {
          name_col:"BJ"
        },
         {
          name_col:"BK"
        },
         {
          name_col:"BL"
        },
         {
          name_col:"BM"
        },
         {
          name_col:"BN"
        },
         {
          name_col:"BO"
        },
         {
          name_col:"BP"
        },
         {
          name_col:"BQ"
        },
         {
          name_col:"BR"
        },
         {
          name_col:"BS"
        },
         {
          name_col:"BT"
        },
         {
          name_col:"BU"
        },
         {
          name_col:"BV"
        },
         {
          name_col:"BW"
        },
       
        
      ],

      column_history:[
          {
            name_col:"B"
          },
           {
            name_col:"C"
          },
           {
            name_col:"D"
          },
           {
            name_col:"E"
          },
          {
            name_col:"F"
          }
      ],
      phase_type:[
        {
          per:0.29
        },
        {
          per:0.35
        },
        {
          per:0.51
        },
        {
          per:0.57
        },
      ],
      count_work_week:[
        {
          p:44
        },
        {
          p:45
        },
        {
          p:46
        },
        {
          p:47
        },
        {
          p:48
        },
        {
          p:49
        },
        {
          p:50
        },
        {
          p:51
        },
        {
          p:52
        },
        {
          p:1
        },
        {
          p:2
        },
        {
          p:3
        },
        {
          p:4
        },
        {
          p:5
        },
        {
          p:6
        },
        {
          p:7
        },
        {
          p:8
        },
        {
          p:9
        },
        {
          p:10
        },
        {
          p:11
        },
        {
          p:12
        },
        {
          p:13
        },
        {
          p:14
        },
        {
          p:15
        },
        {
          p:16
        },
        {
          p:17
        },
        {
          p:18
        },
        {
          p:19
        },
        {
          p:20
        },
        {
          p:21
        },
        {
          p:22
        },
        {
          p:23
        },
        {
          p:24
        },
        {
          p:25
        },
        {
          p:26
        },
        {
          p:27
        },
         {
          p:28
        },
        {
          p:29
        },
        {
          p:30
        },
        {
          p:31
        },
        {
          p:32
        },
        {
          p:33
        },
         {
          p:34
        },
        {
          p:35
        },
        {
          p:36
        },
        {
          p:37
        },
        {
          p:38
        },
        {
          p:39
        },
        {
          p:40
        },
        {
          p:42
        },
        {
          p:43
        },
        {
          p:44
        },

      ]


    }
  },
async  mounted() {
  this.year=[]
      this.end_loss_value=[]
      this.splice_loss_value=[]
      this.cut_out_loss_value=[]
      this.remnants_loss_value=[]
      this.total_spread_loss_per_value=[]
      this.total_width_loss_per_value=[]
    let org = this.$q.localStorage.getItem("org");
    this.org = org
  await  axios.get(this.$api_url + "/mainso.php/get_target")
      .then((resp)=>{
      
        this.rowx = resp.data.data

  
        this.end_loss_woven = parseFloat(this.rowx[0].END_LOSS_WOVEN).toFixed(2)
        this.end_loss_light = parseFloat(this.rowx[0].END_LOSS_LIGHT).toFixed(2)
        this.end_loss_fleece = parseFloat(this.rowx[0].END_LOSS_FLEECE).toFixed(2)
        this.end_loss_special = parseFloat(this.rowx[0].END_LOSS_SPECIAL).toFixed(2)

        this.splice_loss_woven = parseFloat(this.rowx[0].SPLICE_LOSS_WOVEN).toFixed(2)
        this.splice_loss_light = parseFloat(this.rowx[0].SPLICE_LOSS_LIGHT).toFixed(2)
        this.splice_loss_fleece = parseFloat(this.rowx[0].SPLICE_LOSS_FLEECE).toFixed(2)
        this.splice_loss_special = parseFloat(this.rowx[0].SPLICE_LOSS_SPECIAL).toFixed(2)

        this.cut_loss_woven = parseFloat(this.rowx[0].CUT_LOSS_WOVEN).toFixed(2)
        this.cut_loss_light = parseFloat(this.rowx[0].CUT_LOSS_LIGHT).toFixed(2)
        this.cut_loss_fleece = parseFloat(this.rowx[0].CUT_LOSS_FLEECE).toFixed(2)
        this.cut_loss_special = parseFloat(this.rowx[0].CUT_LOSS_SPECIAL).toFixed(2)

        this.rem_loss_woven = parseFloat(this.rowx[0].REM_LOSS_WOVEN).toFixed(2)
        this.rem_loss_light = parseFloat(this.rowx[0].REM_LOSS_LIGHT).toFixed(2)
        this.rem_loss_fleece = parseFloat(this.rowx[0].REM_LOSS_FLEECE).toFixed(2)
        this.rem_loss_special = parseFloat(this.rowx[0].REM_LOSS_SPECIAL).toFixed(2) 

        this.total_woven = Number(this.end_loss_woven) + Number(this.splice_loss_woven) + Number(this.cut_loss_woven)
        +Number(this.rem_loss_woven)

        this.total_light = Number(this.end_loss_light) + Number(this.splice_loss_light) + Number(this.cut_loss_light)
        +Number(this.rem_loss_light)

        this.total_fleece = Number(this.end_loss_fleece) + Number(this.splice_loss_fleece) + Number(this.cut_loss_fleece)
        +Number(this.rem_loss_fleece)

        this.total_special = Number(this.end_loss_special) + Number(this.splice_loss_special) + Number(this.cut_loss_special)
        +Number(this.rem_loss_special)

        this.total_woven = parseFloat(this.total_woven).toFixed(2)
        this.total_light = parseFloat(this.total_light).toFixed(2)
        this.total_fleece = parseFloat(this.total_fleece).toFixed(2)
        this.total_special = parseFloat(this.total_special).toFixed(2)
        
      }).catch((error)=>{
        console.log(error)
      })
      
       /* function convert_date(val){
          let day = val.substring(0, 2);
          let month = val.substring(3, 5);
          let year = val.substring(6, 10);
          this.start_date = year +"/"+ month +"/"+ day
          return this.start_date
          
        }
 */
    await  axios.get(this.$api_url + "/mainso.php/get_date2")
      .then((resp)=>{
        this.first_day = (resp.data.data[0].START_DATE)
      }).catch((error)=>{
        console.log(error)
      })

      await axios.get(this.$api_url + "/mainso.php/get_date2")
      .then((resp)=>{
        this.first_day = (resp.data.data[0].START_DATE)
      }).catch((error)=>{
        console.log(error)
      })

      this.keep_history_data = []
      const params2 = new FormData();
      params2.append("org",this.org)
     await axios({
        method:'post',
        url: this.$api_url + '/mainso.php/find_history_data',
        data:params2,
      }).then((resp)=>{
        console.log(resp.data)
        this.keep_history_data = resp.data.data
      })
      
      console.log(this.keep_history_data)
      for(var ax = 0; ax<this.keep_history_data.length; ax++){
        this.year[ax] = this.keep_history_data[ax].YEAR
        this.end_loss_value[ax] = this.keep_history_data[ax].END_LOSS
        this.splice_loss_value[ax] = this.keep_history_data[ax].SPLICE_LOSS
        this.cut_out_loss_value[ax] = this.keep_history_data[ax].CUT_OUT_LOSS
        this.remnants_loss_value[ax] = this.keep_history_data[ax].REMNANTS_LOSS
        this.total_spread_loss_per_value[ax] = this.keep_history_data[ax].TOTAL_SPREAD_LOSS_PER
        this.total_width_loss_per_value[ax] = this.keep_history_data[ax].TOTAL_WIDTH_LOSS_PER
      }
      console.log(this.year)

  },
methods: {
    async exportexcel() {
    try{
       this.$q.loading.show({
          spinner: QSpinnerGears,
          spinnerColor: 'wthite',
          spinnerSize: 180,
          backgroundColor: 'black',
         
        })
      this.total_audit = []
      this.total_target=[]
      this.ac_end_loss=[]
      this.ac_splice_loss=[]
      this.ac_cut_out_loss=[]
      this.ac_remnants_loss=[]
      this.ac_end_sum=[]
      this.ac_splice_sum=[]
      this.ac_cut_sum=[]
      this.ac_rem_sum=[]
      this.all_find_max=[]
      this.last_total_spread_end=[]
      this.last_total_spread_splice=[]
      this.last_total_spread_cut=[]
      this.last_total_spread_rem=[]
       this.total_spread_loss_woven_data=[]
      this.total_spread_loss_woven_target=[]
      this.total_spread_loss_light_data=[]
      this.total_spread_loss_light_target=[]
      this.total_spread_loss_fleece_data=[]
      this.total_spread_loss_fleece_target=[]
      this.total_spread_loss_special_data=[]
      this.total_spread_loss_special_target=[]
      this.arr_end_loss_1_4=[]
      this.date_data_for_graph1=[]
      this.date_data_for_graph2=[]
      this.date_data_for_graph3=[]
      this.date_data_for_graph4=[]

      this.data_graph1=[]
      this.data_graph2=[]
      this.data_graph3=[]
      this.data_graph4=[]

      this.sum_all_4_ttl_marker=[]

      this.data_target_graph1=[]
      this.data_target_graph2=[]
      this.data_target_graph3=[]
      this.data_target_graph4=[]

      this.total_spread_loss_woven_data_find_0 = []

      this.total_spread_loss_light_data_find_0 = []

      this.total_spread_loss_fleece_data_find_0 = []

      this.total_spread_loss_special_data_find_0 = []

      this.sum_widthloss=[]
      this.sum_ttl_marker=[]

      this.sum_width_loss_ttl_marker = []

      this.target_width=[]
      this.acc_width_arr=[]
      this.end_loss_target_arr=[]
      this.count_result_arr=[]

      this.result_per_week=[]
 	    this.result_max_woven=[]
      this.result_max_light=[]
      this.result_max_fleece=[]
      this.result_max_special=[]

      


      
      
         const workbook = new Excel.Workbook();
      workbook.creator = "Nanyang";
      workbook.lastModifiedBy = "Admin";
      workbook.created = new Date(2021, 8, 30);
      workbook.modified = new Date();
      workbook.lastPrinted = new Date(2021, 7, 27);
      const worksheet = workbook.addWorksheet("Data");
     

      worksheet.columns = [
        { key: "A", width: 10 },
        { key: "B", width: 12 },
        { key: "C", width: 7.63 },
        { key: "D", width: 12 },
        { key: "E", width: 10 },
        { key: "F", width: 15 },
        { key: "G", width: 15 },
        { key: "H", width: 15 },
        { key: "I", width: 15 },
        { key: "J", width: 12 },
        { key: "K", width: 12 },
        { key: "L", width: 15 },
        { key: "M", width: 15 },
        { key: "N", width: 6.7 },
        { key: "O", width: 6.7 },
        { key: "P", width: 6.7 },
        { key: "Q", width: 6.7 },
        { key: "R", width: 6.7 },
        { key: "S", width: 12 },
        { key: "T", width: 12 },
        { key: "U", width: 6.7 },
        { key: "V", width: 6.7 },
        { key: "W", width: 22.7 },
        { key: "X", width: 12 },
        { key: "AC", width: 15.13 },
      ];

      //setting cell and border

for(var ax=0; ax<this.arr_left_colume.length; ax++){
    
      for(var az=1; az<14; az++){
    worksheet.getCell(this.arr_left_colume[ax].name_col+az).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
        worksheet.getCell(this.arr_left_colume[ax].name_col+az).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },  
          right: { style: "thin" },
        };
       worksheet.getCell(this.arr_left_colume[ax].name_col+az).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
      }
 }
      for(var ax=0; ax<this.arr_colume.length; ax++){
    
      for(var az=1; az<160; az++){
    worksheet.getCell(this.arr_colume[ax].name_col+az).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
       worksheet.getCell(this.arr_colume[ax].name_col+az).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
      }
 }


      for(var az=1; az<160; az++){
    worksheet.getCell("BX"+az).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
       worksheet.getCell("BX"+az).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
      }
 

 
    
      for(var az=1; az<160; az++){
    worksheet.getCell("W"+az).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: true
      };
       worksheet.getCell("W"+az).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
      }

 
    
     
      
    
     
   
 
        worksheet.getCell("A3").value = "Descriptions"; 
       worksheet.mergeCells("A3:A5");
     
      
      
       worksheet.getCell("B3").value = "History Performance"; 
       worksheet.mergeCells("B3:F3");
     
     
  
      worksheet.getCell("B4").value = this.year[0]; 
       worksheet.mergeCells("B4:B5");
    
       
        

          worksheet.getCell("C4").value = this.year[1]; 
       worksheet.mergeCells("C4:C5");
      
    

      worksheet.getCell("D4").value = this.year[2]; 
       worksheet.mergeCells("D4:D5");
      
      

          worksheet.getCell("E4").value = this.year[3]; 
       worksheet.mergeCells("E4:E5");
   
    


           worksheet.getCell("F4").value = this.year[4]; 
       worksheet.mergeCells("F4:F5");
    
      
       
        
        /*   worksheet.getCell("G4").value = "2021 Period ( 26 Oct 2020 - 23 Oct 2021)"; 
       worksheet.mergeCells("G4:G5"); */
  
      

        worksheet.getCell("G3").value = "Phase 6 Target(Current)"; 
       worksheet.mergeCells("G3:J3");
    
     
        
        worksheet.getCell("K3").value = "Weighted Target of This Week"; 
       worksheet.mergeCells("K3:K4");
    
   

          worksheet.getCell("G4").value = "Woven"; 
       worksheet.mergeCells("G4:G5");
  
     
       

          worksheet.getCell("H4").value = "Knit"; 
       worksheet.mergeCells("H4:J4");
    
    
        worksheet.getCell("H5").value = "Ligth Middle Weight"; 
       
    
    
    

         worksheet.getCell("I5").value = "Fleece"; 
       
     
  

          worksheet.getCell("J5").value = "Special"; 
  
     

          worksheet.getCell("K5").value = "Overall"; 
    
     
      
 
            //Actual Performance
 
            worksheet.getCell("L3").value = "Actual Performance"; 
       worksheet.mergeCells("L3:Q3");
   
     
    

            worksheet.getCell("L4").value = "Woven"; 
       worksheet.mergeCells("L4:L5");
      
     

      worksheet.getCell("M4").value = "Knit"; 
       worksheet.mergeCells("M4:O4");
     
      
    

         worksheet.getCell("M5").value = "Ligth Middle Weight"; 
       
     
     
      
         worksheet.getCell("N5").value = "Fleece"; 
       
    
     


          worksheet.getCell("O5").value = "Special"; 
    
     
    
        worksheet.getCell("P4").value = "Weighted  Actual of This Week"; 
       worksheet.mergeCells("P4:Q4");
     
     
        worksheet.getCell("R5").value = "Average"; 
   
      
   
        worksheet.getCell("S5").value = "Total"; 
    
    
      this.cumu_since = moment(this.first_day).format("DD-MM-YYYY");
         worksheet.getCell("R3").value = "Cumulated since" + this.cumu_since; 
       worksheet.mergeCells("R3:S4");
    
    

          worksheet.getCell("R5").value = "Average"; 
  
   
    
      

        worksheet.getCell("S5").value = "Total"; 
     
    
    

        
        worksheet.getCell("A6").value = "Total  Marker Audit  (Yds)"; 
   
     
         worksheet.getCell("A7").value = "End Loss"; 
     
      
       
         worksheet.getCell("A8").value = "Splice Loss"; 
       
      
    

         worksheet.getCell("A9").value = "Cut out Loss"; 
   
      
     

          worksheet.getCell("A10").value = "Remnants Loss"; 
     
       
        
       worksheet.getCell("A11").value = "Total Spread Loss audit"; 
      

       
       for(var axc = 0; axc<this.column_history.length; axc++){
        worksheet.getCell(this.column_history[axc].name_col+7).value = this.end_loss_value[axc]/100
        worksheet.getCell(this.column_history[axc].name_col+7).numFmt = "0.00%"

        worksheet.getCell(this.column_history[axc].name_col+8).value = this.splice_loss_value[axc]/100
        worksheet.getCell(this.column_history[axc].name_col+8).numFmt = "0.00%"

        worksheet.getCell(this.column_history[axc].name_col+9).value = this.cut_out_loss_value[axc]/100
        worksheet.getCell(this.column_history[axc].name_col+9).numFmt = "0.00%"

        worksheet.getCell(this.column_history[axc].name_col+10).value = this.remnants_loss_value[axc]/100
        worksheet.getCell(this.column_history[axc].name_col+10).numFmt = "0.00%"

        worksheet.getCell(this.column_history[axc].name_col+11).value = this.total_spread_loss_per_value[axc]/100
        worksheet.getCell(this.column_history[axc].name_col+11).numFmt = "0.00%"


      }


      
      const params  = new FormData();
      params.append("start",this.start_date)
      params.append("end",this.end_date)
       let org = this.$q.localStorage.getItem("org");
       
      params.append("org",org)
      for (var pair of params.entries()) {
            console.log(pair[0] + ", " + pair[1]);
      }
          

      await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data_weekly',
        data:params,
      }).then((resp)=>{
       
        this.rowexport =[]
        
        function check_null(val){
                if(val == '' ||  val == null || val == "NaN" || val == "null" || val == undefined ){
                  return 0
                }else{
                  return val
                }
        }
       
        if(resp.data.data.length > 0){
             resp.data.data.forEach((e) => {
                this.rowexport.push({
                  mu_date : e.MU_DATE,
                  gm_table:e.GM_TABLE,
                  so_no_doc:e.SO_NO_DOC,
                  customer:e.CUSTOMER_NAME,
                  gm_no:e.GM_NO,
                  gm_no_short :e.GM_NO.slice(2,6),
                  fabric_type:e.FABRIC_TYPE,
                  f_type:e.F_TYPE,
                  issue: e.ISSUE,
                  issue_roll :e.ROLL_NO,
                  avg_width:e.AVG_WIDTH,
                  gmm2: e.WEIGHT_M2,
                  gmyd: e.WEIGHT_YD,
                  width_inch : e.WIDTH_INCH,
                  length_ydb:e.YARD,
                  length_inch: e.LENGTH_INCH,
                  obs_width: e.AVG_WIDTH,
                  marker: e.MARKER,
                  roll:e.ROLL_NO,
                  ply: e.PLY,
                  endplug1: e.END_PLUG1,
                  endplug2: e.END_PLUG2,
                  cause: e.CAUSE_NAME,
                  code: e.CAUSE_CODE,
                  splicenoof: e.SPLICE,
                  spliceinch: e.SPLICE_INCH,
                  cutoutnoof: e.CUT_OUT,
                  cutoutweight: e.CUT_OUT_GRAM,
                  remnants: check_null(e.REMNANTS_WEIGHT),
                  widthplug1: e.PLUG1_GRAM,
                  widthplug2: e.PLUG2_GRAM,
                  avgglueside: e.GLUE_SIDE,
                  ttl_marker: parseFloat(e.TTL_MARKER_YD).toFixed(2),
                  yard_roll:e.YARD_ROLL,
                  length_yd:e.LENGTH_YD,
                  obswidth:e.OBS_WIDTH,
                  ttlobs_widthloss:e.OBS_WIDTH,
                  end1_2:(e.END1_2),
                  plug1_inch:e.PLUG1_INCH,
                  plug2_inch:e.PLUG2_INCH,
                  per_plug1:e.PER_PLUG1,
                  per_plug2:e.PER_PLUG2,
                  plug12:e.PLUG1_2_INCH,
                  endloss_yd: e.END_LOSS_YD,
                  widthloss:e.WITH_LOSS,
                  endless_length_yd:e.AVG_END_LOSS_YD,
                  avg_end:e.AVG_END_LOSS,
                  sectionlosslenth:check_null(e.SECTION_LOSS_YD),
                  sectionlossper:check_null(e.SECTION_LOSS_PER),
                  spliceloss:check_null(e.SPLICE_LOSS_YD),
                  splicelossper:parseFloat(check_null(e.SPLICE_LOSS_PER)).toFixed(2),
                  overlap:e.OVER_LAB_YD,
                  overlapper:parseFloat(e.OVER_LAB_PER).toFixed(2),
                  defect:parseFloat(e.DEFECT_YD).toFixed(2),
                  defectper:parseFloat(e.DEFECT_PER).toFixed(2),
                  totalcutout:parseFloat(check_null(e.TOTAL_CUTOUT_YD)).toFixed(2),
                  totalcutoutper:check_null(e.TOTAL_CUTOUT_PER),
                  rement:parseFloat(check_null(e.REMENT_LOSS_YD)).toFixed(2),
                  rement_per:parseFloat(check_null(e.REMENT_LOSS_PER)).toFixed(2),
                  totallenthloss:parseFloat(check_null(e.TOTAL_LENGTH_LOSS_YD)).toFixed(2),
                  totallenthlossper:parseFloat(check_null(e.TOTAL_LEN_LOSS_PER)).toFixed(2),
                  avg_end_code:e.AVG_END_CODE,
                  avg_end_cause:e.AVG_END_CAUSE,
                  per_splice_loss_code:e.SPLICE_LOSS_PER_CODE,
                  per_splice_loss_cause:e.SPLICE_LOSS_PER_CAUSE,
                  remnants_loss_code:e.REMNANTS_LOSS_CODE,
                  remnants_loss_cause:e.REMNANTS_LOSS_CAUSE,
                  total_cut_out_per_code:e.TOTAL_CUT_OUT_PER_CODE,
                  total_cut_out_per_cause:e.TOTAL_CUT_OUT_PER_CAUSE,
                  widthloss_code:e.WIDTH_LOSS_CODE,
                  widthloss_cause:e.WIDTH_LOSS_CAUSE,
                  balance: parseFloat(e.ISSUE-((e.TTL_MARKER_YD*e.WEIGHT_YD) /1000)).toFixed(2),
                })
               
       });
        }
        this.all_result_ttlmarker_woven = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "1"){
              this.all_result_ttlmarker_woven = Number(this.all_result_ttlmarker_woven) + Number(this.rowexport[ax].ttl_marker)
            }
          }

          

           worksheet.getCell("L6").value = Number(this.all_result_ttlmarker_woven);
           worksheet.getCell("L6").numFmt = "#,##0.00"

           
    
       
          
       
        this.all_result_ttlmarker_light = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "2"){
              this.all_result_ttlmarker_light = Number(this.all_result_ttlmarker_light) + Number(this.rowexport[ax].ttl_marker)
            }
          }
         
         console.log(this.all_result_ttlmarker_light)
        
       worksheet.getCell("M6").value = Number(this.all_result_ttlmarker_light); 
       worksheet.getCell("M6").numFmt = "#,##0.00"

     
      
       
          this.all_result_ttlmarker_fleece = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "3"){
              this.all_result_ttlmarker_fleece = Number(this.all_result_ttlmarker_fleece) + Number(this.rowexport[ax].ttl_marker)
            }
          }
          worksheet.getCell("N6").value = Number(this.all_result_ttlmarker_fleece);
          worksheet.getCell("N6").numFmt = "#,##0.00"
    
      

          this.all_result_ttlmarker_special = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "4"){
              this.all_result_ttlmarker_special = Number(this.all_result_ttlmarker_special) + Number(this.rowexport[ax].ttl_marker)
            }
          }
            worksheet.getCell("O6").value = Number(this.all_result_ttlmarker_special);
            worksheet.getCell("O6").numFmt = "#,##0.00"
      
        

          this.all_type_summary_ttlmarker = Number(this.all_result_ttlmarker_woven)+Number(this.all_result_ttlmarker_light)+Number(this.all_result_ttlmarker_fleece)
          +Number(this.all_result_ttlmarker_special)

         
       worksheet.getCell("P6").value = Number(this.all_type_summary_ttlmarker);
       worksheet.getCell("P6").numFmt = "#,##0.00"
      
       
      
        this.azc = 7
        for(var ax= 0; ax<5; ax++){
          this.azc = 7+ax
          
          worksheet.getCell("Q"+this.azc).value = ""; 
     
      
       
        } 

        worksheet.getCell("Q13").value = ""; 
       
      
        
        this.azc4 = 12
        for(var ax= 0; ax<2; ax++){
          this.azc4 = 12+ax
          
          worksheet.getCell("R"+this.azc4).value = ""; 
     
       
        
        } 

        
        
          
          worksheet.getCell("R6").value = ""; 
     
       
        
         

          this.azc3 = 7
        for(var ax= 0; ax<6; ax++){
          this.azc3 = 7+ax
          
          worksheet.getCell("S"+this.azc3).value = ""; 
       
       
        
        } 
          
      })
      this.Woven_total = 0
      this.Woven_total = Number(this.end_loss_woven)+Number(this.splice_loss_woven)+Number(this.cut_loss_woven)+Number(this.rem_loss_woven)
      this.total_audit.push({
        value_total:this.Woven_total
      })
       worksheet.getCell("G7").value = this.end_loss_woven/100
       worksheet.getCell("G7").numFmt = '0.00%'; 
      
       
         worksheet.getCell("G8").value = this.splice_loss_woven/100
         worksheet.getCell("G8").numFmt = '0.00%'; 
     
    
      
         worksheet.getCell("G9").value = this.cut_loss_woven/100
         worksheet.getCell("G9").numFmt = '0.00%'; 
      
     
         worksheet.getCell("G10").value = this.rem_loss_woven/100
         worksheet.getCell("G10").numFmt = '0.00%'; 
      
      
       
         worksheet.getCell("G11").value =this.Woven_total/100
         worksheet.getCell("G11").numFmt = '0.00%'; 
       
      

      this.Light_total = 0
      this.Light_total = Number(this.end_loss_light)+Number(this.splice_loss_light)+Number(this.cut_loss_light)+Number(this.rem_loss_light)
     this.total_audit.push({
        value_total:this.Light_total
      })
       worksheet.getCell("H7").value =this.end_loss_light/100
       worksheet.getCell("H7").numFmt = '0.00%'; 
      
        
       
         worksheet.getCell("H8").value = this.splice_loss_light/100
         worksheet.getCell("H8").numFmt = '0.00%'; 
      
        
       
         worksheet.getCell("H9").value = this.cut_loss_light/100
         worksheet.getCell("H9").numFmt = '0.00%'; 
      
      
        
         worksheet.getCell("H10").value = this.rem_loss_light/100
         worksheet.getCell("H10").numFmt = '0.00%'; 
      
        
       
         worksheet.getCell("H11").value = this.Light_total/100
         worksheet.getCell("H11").numFmt = '0.00%'; 
       
        
        

        this.Fleece_total = 0
      this.Fleece_total = Number(this.end_loss_fleece)+Number(this.splice_loss_fleece)+Number(this.cut_loss_fleece)+Number(this.rem_loss_fleece)
     this.total_audit.push({
        value_total:this.Fleece_total
      })
       worksheet.getCell("I7").value = this.end_loss_fleece/100
       worksheet.getCell("I7").numFmt = '0.00%'; 
      
        
       
         worksheet.getCell("I8").value = this.splice_loss_fleece/100
         worksheet.getCell("I8").numFmt = '0.00%'; 
      
        
        
         worksheet.getCell("I9").value = this.cut_loss_fleece/100
         worksheet.getCell("I9").numFmt = '0.00%'; 
      
       
        
         worksheet.getCell("I10").value = this.rem_loss_fleece/100
         worksheet.getCell("I10").numFmt = '0.00%'; 
       
       
        
         worksheet.getCell("I11").value = this.Fleece_total/100
         worksheet.getCell("I11").numFmt = '0.00%'; 
       
       
       

      this.Special_total = 0
      this.Special_total = Number(this.end_loss_special)+Number(this.splice_loss_special)+Number(this.cut_loss_special)+Number(this.rem_loss_special)
     this.total_audit.push({
        value_total:this.Special_total
      })
       worksheet.getCell("J7").value = this.end_loss_special/100
       worksheet.getCell("J7").numFmt = '0.00%'; 
      
        
       
         worksheet.getCell("J8").value = this.splice_loss_special/100
         worksheet.getCell("J8").numFmt = '0.00%'; 
     
        
       
         worksheet.getCell("J9").value = this.cut_loss_special/100
         worksheet.getCell("J9").numFmt = '0.00%'; 
      
       
       
         worksheet.getCell("J10").value = this.rem_loss_special/100
         worksheet.getCell("J10").numFmt = '0.00%'; 
     
       
        
         worksheet.getCell("J11").value = this.Special_total/100
         worksheet.getCell("J11").numFmt = '0.00%'; 
     
       
       

        //endloss
          this.actual_per_woven_end = this.end_loss_woven * this.all_result_ttlmarker_woven
     
          this.actual_per_light_end = this.end_loss_light * this.all_result_ttlmarker_light

          this.actual_per_fleece_end = this.end_loss_fleece * this.all_result_ttlmarker_fleece
    
          this.actual_per_special_end = this.end_loss_special * this.all_result_ttlmarker_special
        
          
          this.all_actual_endloss = Number(this.actual_per_woven_end) + Number(this.actual_per_light_end) + Number(this.actual_per_fleece_end) + Number(this.actual_per_special_end)
          
          this.all_actual_endloss = this.all_actual_endloss / this.all_type_summary_ttlmarker
   

            worksheet.getCell("K7").value = this.all_actual_endloss/100
            worksheet.getCell("K7").numFmt = '0.00%'; 
      
      
      //light

          this.actual_per_woven_splice = this.splice_loss_woven * this.all_result_ttlmarker_woven
         
          this.actual_per_light_splice = this.splice_loss_light * this.all_result_ttlmarker_light
         
          this.actual_per_fleece_splice = this.splice_loss_fleece * this.all_result_ttlmarker_fleece
      
          this.actual_per_special_splice = this.splice_loss_special * this.all_result_ttlmarker_special
        
          
          this.all_actual_spliceloss = Number(this.actual_per_woven_splice) + Number(this.actual_per_light_splice) + Number(this.actual_per_fleece_splice) + Number(this.actual_per_special_splice)
          
          this.all_actual_spliceloss = this.all_actual_spliceloss / this.all_type_summary_ttlmarker
     



            worksheet.getCell("K8").value = this.all_actual_spliceloss/100
            worksheet.getCell("K8").numFmt = '0.00%'; 
     
           //cutoutloss

          this.actual_per_woven_cutoutloss = this.cut_loss_woven * this.all_result_ttlmarker_woven
     
          this.actual_per_light_cutoutloss = this.cut_loss_light * this.all_result_ttlmarker_light

          this.actual_per_fleece_cutoutloss = this.cut_loss_fleece * this.all_result_ttlmarker_fleece
    
          this.actual_per_special_cutoutloss = this.cut_loss_special * this.all_result_ttlmarker_special
        
          
          this.all_actual_cutoutloss = Number(this.actual_per_woven_cutoutloss) + Number(this.actual_per_light_cutoutloss) + Number(this.actual_per_fleece_cutoutloss) + Number(this.actual_per_special_cutoutloss)
          
          this.all_actual_cutoutloss = this.all_actual_cutoutloss / this.all_type_summary_ttlmarker
   


            worksheet.getCell("K9").value = this.all_actual_cutoutloss/100
            worksheet.getCell("K9").numFmt = '0.00%'; 
     


          //remenants

          this.actual_per_woven_remenants = this.rem_loss_woven * this.all_result_ttlmarker_woven
     
          this.actual_per_light_remenants = this.rem_loss_light * this.all_result_ttlmarker_light

          this.actual_per_fleece_remenants = this.rem_loss_fleece * this.all_result_ttlmarker_fleece
    
          this.actual_per_special_remenants = this.rem_loss_special * this.all_result_ttlmarker_special
        
          
          this.all_actual_remnantsloss = Number(this.actual_per_woven_remenants) + Number(this.actual_per_light_remenants) + Number(this.actual_per_fleece_remenants) + Number(this.actual_per_special_remenants)
          
          this.all_actual_remnantsloss = this.all_actual_remnantsloss / this.all_type_summary_ttlmarker
        


        
            worksheet.getCell("K10").value = this.all_actual_remnantsloss/100
            worksheet.getCell("K10").numFmt = '0.00%'; 
      

        this.overall_total = Number(this.all_actual_endloss) + Number(this.all_actual_spliceloss)+ Number(this.all_actual_cutoutloss)
        + Number(this.all_actual_remnantsloss)
      
         worksheet.getCell("K11").value = this.overall_total/100
         worksheet.getCell("K11").numFmt = '0.00%'; 
    
      
      this.actual_woven_endloss = 0
        for(var axa = 0; axa<this.rowexport.length; axa++){
          if(this.rowexport[axa].endless_length_yd > 0){
            if(this.rowexport[axa].f_type == "1")
            this.actual_woven_endloss = Number(this.actual_woven_endloss) + Number(this.rowexport[axa].endless_length_yd)
          }
        }
        
        this.actual_knit_woven = 0
        if(this.actual_woven_endloss > 0){
        this.actual_knit_woven = this.actual_woven_endloss / this.all_result_ttlmarker_woven * 100
        }
        if(this.actual_knit_woven == 0){
         this.ac_end_loss.push({
          value_end :"-",
          value_name:"end_loss_woven"
        })
        }else{
           this.ac_end_loss.push({
          value_end :this.actual_knit_woven,
          value_name:"end_loss_woven"
        })
        }
              
          worksheet.getCell("L7").value = this.actual_knit_woven/100
          worksheet.getCell("L7").numFmt = '0.00%'; 
    

        this.actual_light_endloss = 0
        for(var axc = 0; axc<this.rowexport.length; axc++){
          if(this.rowexport[axc].endless_length_yd > 0){
            if(this.rowexport[axc].f_type == "2")
            this.actual_light_endloss = Number(this.actual_light_endloss) + Number(this.rowexport[axc].endless_length_yd)
          }
        }
        this.actual_knit_light = 0

        if(this.actual_light_endloss > 0){
        this.actual_knit_light = this.actual_light_endloss / this.all_result_ttlmarker_light * 100
        }
        if(this.actual_knit_light == 0){
        this.ac_end_loss.push({
          value_end :"-",
          value_name:"end_loss_light"
        })
        }else{
          this.ac_end_loss.push({
          value_end :this.actual_knit_light,
          value_name:"end_loss_light"
        })
        }
          worksheet.getCell("M7").value = this.actual_knit_light/100
          worksheet.getCell("M7").numFmt = '0.00%'; 
      

           this.actual_fleece_endloss = 0
        for(var axc = 0; axc<this.rowexport.length; axc++){
          if(this.rowexport[axc].endless_length_yd > 0){
            if(this.rowexport[axc].f_type == "3")
            this.actual_fleece_endloss = Number(this.actual_fleece_endloss) + Number(this.rowexport[axc].endless_length_yd)
          }
        }
        this.actual_knit_fleece = 0
        if(this.actual_fleece_endloss > 0){
        this.actual_knit_fleece = this.actual_fleece_endloss / this.all_result_ttlmarker_fleece * 100
        }
        if(this.actual_knit_fleece == 0){
        this.ac_end_loss.push({
          value_end :"-",
          value_name:"end_loss_fleece"
        })
        }else{
          this.ac_end_loss.push({
          value_end :this.actual_knit_fleece,
          value_name:"end_loss_fleece"
        })
        }
              
          worksheet.getCell("N7").value = this.actual_knit_fleece/100
          worksheet.getCell("N7").numFmt = '0.00%'; 

       

           this.actual_special_endloss = 0
        for(var axc = 0; axc<this.rowexport.length; axc++){
          if(this.rowexport[axc].endless_length_yd > 0){
            if(this.rowexport[axc].f_type == "4")
            this.actual_special_endloss = Number(this.actual_special_endloss) + Number(this.rowexport[axc].endless_length_yd)
          }
        }
        this.actual_knit_special = 0
        if(this.actual_special_endloss > 0){
        this.actual_knit_special = this.actual_special_endloss / this.all_result_ttlmarker_special * 100
        }
        if(this.actual_knit_special == 0){
        this.ac_end_loss.push({
          value_end :"-",
          value_name:"end_loss_special"
        })
        }else{
          this.ac_end_loss.push({
          value_end :this.actual_knit_special,
          value_name:"end_loss_special"
        })
        }

          worksheet.getCell("O7").value = this.actual_knit_special/100
          worksheet.getCell("O7").numFmt = '0.00%'; 
      


         
      this.actual_woven_splice = 0
        for(var axa = 0; axa<this.rowexport.length; axa++){
          if(this.rowexport[axa].endless_length_yd > 0){
            if(this.rowexport[axa].f_type == "1")
            this.actual_woven_splice = Number(this.actual_woven_splice) + Number(this.rowexport[axa].spliceloss)
          }
        }
        this.actual_knit_woven_splice = 0
        if(this.actual_woven_splice > 0){
        this.actual_knit_woven_splice = this.actual_woven_splice / this.all_result_ttlmarker_woven * 100
        }
        if(this.actual_knit_woven_splice == 0){
        this.ac_splice_loss.push({
          value_splice :"-",
          value_name : "splice_woven"
        })
        }else{
        this.ac_splice_loss.push({
          value_splice :this.actual_knit_woven_splice,
          value_name : "splice_woven"
        })
        }
                 
          worksheet.getCell("L8").value = this.actual_knit_woven_splice/100
          worksheet.getCell("L8").numFmt = '0.00%'; 
      


        this.actual_light_splice = 0
        for(var axa = 0; axa<this.rowexport.length; axa++){
          if(this.rowexport[axa].endless_length_yd > 0){
            if(this.rowexport[axa].f_type == "2")
            this.actual_light_splice = Number(this.actual_light_splice) + Number(this.rowexport[axa].spliceloss)
          }
        }
        this.actual_knit_light_splice = 0
        if(this.actual_light_splice > 0){
        this.actual_knit_light_splice = this.actual_light_splice / this.all_result_ttlmarker_light * 100
        }
      
         if(this.actual_knit_light_splice == 0){
        this.ac_splice_loss.push({
          value_splice :"-",
          value_name : "splice_light"
        })
        }else{
        this.ac_splice_loss.push({
          value_splice :this.actual_knit_light_splice,
          value_name : "splice_light"
        })
        }
          worksheet.getCell("M8").value = this.actual_knit_light_splice/100
          worksheet.getCell("M8").numFmt = '0.00%'; 
    
        this.actual_fleece_splice = 0
        for(var axa = 0; axa<this.rowexport.length; axa++){
          if(this.rowexport[axa].endless_length_yd > 0){
            if(this.rowexport[axa].f_type == "3")
            this.actual_fleece_splice = Number(this.actual_fleece_splice) + Number(this.rowexport[axa].spliceloss)
          }
        }
        this.actual_knit_fleece_splice = 0
        if(this.actual_fleece_splice > 0){
        this.actual_knit_fleece_splice = this.actual_fleece_splice / this.all_result_ttlmarker_fleece * 100
        }
      
          if(this.actual_knit_fleece_splice == 0){
        this.ac_splice_loss.push({
          value_splice :"-",
          value_name : "splice_fleece"
        })
        }else{
        this.ac_splice_loss.push({
          value_splice :this.actual_knit_fleece_splice,
          value_name : "splice_fleece"
        })
        }
          worksheet.getCell("N8").value = this.actual_knit_fleece_splice/100
          worksheet.getCell("N8").numFmt = '0.00%'; 
 

         this.actual_special_splice = 0
        for(var axa = 0; axa<this.rowexport.length; axa++){
          if(this.rowexport[axa].endless_length_yd > 0){
            if(this.rowexport[axa].f_type == "4")
            this.actual_special_splice = Number(this.actual_special_splice) + Number(this.rowexport[axa].spliceloss)
          }
        }
        this.actual_knit_special_splice = 0
        if(this.actual_special_splice > 0){
        this.actual_knit_special_splice = this.actual_special_splice / this.all_result_ttlmarker_special * 100
        }
          if(this.actual_knit_special_splice == 0){
        this.ac_splice_loss.push({
          value_splice :this.actual_knit_special_splice,
          value_name : "splice_special"
        })
        }else{
        this.ac_splice_loss.push({
          value_splice :"-",
          value_name : "splice_special"
        })
        }    
          worksheet.getCell("O8").value = this.actual_knit_special_splice/100
          worksheet.getCell("O8").numFmt = '0.00%'; 

            
      this.actual_woven_cut = 0
        for(var axa = 0; axa<this.rowexport.length; axa++){
          if(this.rowexport[axa].endless_length_yd > 0){
            if(this.rowexport[axa].f_type == "1")
            this.actual_woven_cut = Number(this.actual_woven_cut) + Number(this.rowexport[axa].totalcutout)
          }
        }
        this.actual_knit_woven_cut = 0
        if(this.actual_woven_cut > 0){
        this.actual_knit_woven_cut = this.actual_woven_cut / this.all_result_ttlmarker_woven * 100
        }
        if(this.actual_knit_woven_cut == 0){
        this.ac_cut_out_loss.push({
          value_cut :"-",
          value_name: "cut_woven"
        })
        }else{
          this.ac_cut_out_loss.push({
          value_cut :this.actual_knit_woven_cut,
          value_name: "cut_woven"
        })
        }
          worksheet.getCell("L9").value = this.actual_knit_woven_cut/100
          worksheet.getCell("L9").numFmt = '0.00%'; 
      

        this.actual_light_cut = 0
        for(var axa = 0; axa<this.rowexport.length; axa++){
          if(this.rowexport[axa].endless_length_yd > 0){
            if(this.rowexport[axa].f_type == "2")
            this.actual_light_cut = Number(this.actual_light_cut) + Number(this.rowexport[axa].totalcutout)
          }
        }
        this.actual_knit_light_cut = 0
        if(this.actual_light_cut > 0){
        this.actual_knit_light_cut = this.actual_light_cut / this.all_result_ttlmarker_light * 100
        }
      
         if(this.actual_knit_light_cut == 0){
        this.ac_cut_out_loss.push({
          value_cut :"-",
          value_name: "cut_light"
        })
        }else{
          this.ac_cut_out_loss.push({
          value_cut :this.actual_knit_light_cut,
          value_name: "cut_light"
        })
        }
          worksheet.getCell("M9").value = this.actual_knit_light_cut/100
          worksheet.getCell("M9").numFmt = '0.00%'; 
      

        
        this.actual_fleece_cut = 0
        for(var axa = 0; axa<this.rowexport.length; axa++){
          if(this.rowexport[axa].endless_length_yd > 0){
            if(this.rowexport[axa].f_type == "3")
            this.actual_fleece_cut = Number(this.actual_fleece_cut) + Number(this.rowexport[axa].totalcutout)
          }
        }
        this.actual_knit_fleece_cut = 0
        if(this.actual_fleece_cut > 0){
        this.actual_knit_fleece_cut = this.actual_fleece_cut / this.all_result_ttlmarker_fleece * 100
        }
      
         if(this.actual_knit_fleece_cut == 0){
        this.ac_cut_out_loss.push({
          value_cut :"-",
          value_name: "cut_fleece"
        })
        }else{
          this.ac_cut_out_loss.push({
          value_cut :this.actual_knit_fleece_cut,
          value_name: "cut_fleece"
        })
        }         
          worksheet.getCell("N9").value = this.actual_knit_fleece_cut/100
          worksheet.getCell("N9").numFmt = '0.00%'; 
      


          this.actual_special_cut = 0
        for(var axa = 0; axa<this.rowexport.length; axa++){
          if(this.rowexport[axa].endless_length_yd > 0){
            if(this.rowexport[axa].f_type == "4")
            this.actual_special_cut = Number(this.actual_special_cut) + Number(this.rowexport[axa].totalcutout)
          }
        }
        this.actual_knit_special_cut = 0
        if(this.actual_special_cut > 0){
        this.actual_knit_special_cut = this.actual_special_cut / this.all_result_ttlmarker_special * 100
        }
      
           if(this.actual_knit_special_cut == 0){
        this.ac_cut_out_loss.push({
          value_cut :"-",
          value_name: "cut_special"
        })
        }else{
          this.ac_cut_out_loss.push({
          value_cut :this.actual_knit_special_cut,
          value_name: "cut_special"
        })
        }
          worksheet.getCell("O9").value = this.actual_knit_special_cut/100
          worksheet.getCell("O9").numFmt = '0.00%'; 
     


        this.actual_woven_rem = 0
        for(var axa = 0; axa<this.rowexport.length; axa++){
          if(this.rowexport[axa].endless_length_yd > 0){
            if(this.rowexport[axa].f_type == "1")
            this.actual_woven_rem = Number(this.actual_woven_rem) + Number(this.rowexport[axa].rement)
          }
        }
        this.actual_knit_woven_rem = 0
        if(this.actual_woven_rem > 0){
        this.actual_knit_woven_rem = this.actual_woven_rem / this.all_result_ttlmarker_woven * 100
        }
        if(this.actual_knit_woven_rem == 0){
        this.ac_remnants_loss.push({
          value_rem :"-",
          value_name : "rem_woven"
        })
        }else{
           this.ac_remnants_loss.push({
          value_rem :this.actual_knit_woven_rem,
          value_name : "rem_woven"
        })
        }   
          worksheet.getCell("L10").value = this.actual_knit_woven_rem/100
          worksheet.getCell("L10").numFmt = '0.00%'; 

       


        this.actual_light_rem = 0
        for(var axa = 0; axa<this.rowexport.length; axa++){
          if(this.rowexport[axa].endless_length_yd > 0){
            if(this.rowexport[axa].f_type == "2")
            this.actual_light_rem = Number(this.actual_light_rem) + Number(this.rowexport[axa].rement)
          }
        }
        this.actual_knit_light_rem = 0
        if(this.actual_light_rem > 0){
        this.actual_knit_light_rem = this.actual_light_rem / this.all_result_ttlmarker_light * 100
        }
      
          if(this.actual_knit_light_rem == 0){
        this.ac_remnants_loss.push({
          value_rem :"-",
          value_name : "rem_light"
        })
        }else{
           this.ac_remnants_loss.push({
          value_rem :this.actual_knit_light_rem,
          value_name : "rem_light"
        })
        }            
          worksheet.getCell("M10").value = this.actual_knit_light_rem/100
          worksheet.getCell("M10").numFmt = '0.00%'; 


        this.actual_fleece_rem = 0
        for(var axa = 0; axa<this.rowexport.length; axa++){
          if(this.rowexport[axa].endless_length_yd > 0){
            if(this.rowexport[axa].f_type == "3")
            this.actual_fleece_rem = Number(this.actual_fleece_rem) + Number(this.rowexport[axa].rement)
          }
        }
        this.actual_knit_fleece_rem = 0
        if(this.actual_fleece_rem > 0){
        this.actual_knit_fleece_rem = this.actual_fleece_rem / this.all_result_ttlmarker_fleece * 100
        }
      
          if(this.actual_knit_fleece_rem == 0){
        this.ac_remnants_loss.push({
          value_rem :"-",
          value_name : "rem_fleece"
        })
        }else{
           this.ac_remnants_loss.push({
          value_rem :this.actual_knit_fleece_rem,
          value_name : "rem_fleece"
        })
        }          
          worksheet.getCell("N10").value = this.actual_knit_fleece_rem/100
          worksheet.getCell("N10").numFmt = '0.00%'; 
   

        this.actual_special_rem = 0
        for(var axa = 0; axa<this.rowexport.length; axa++){
          if(this.rowexport[axa].endless_length_yd > 0){
            if(this.rowexport[axa].f_type == "4")
            this.actual_special_rem = Number(this.actual_special_rem) + Number(this.rowexport[axa].rement)
          }
        }
        this.actual_knit_special_rem = 0
        if(this.actual_special_rem > 0){
        this.actual_knit_special_rem = this.actual_special_rem / this.all_result_ttlmarker_special * 100
        }
      
         if(this.actual_knit_special_rem == 0){
        this.ac_remnants_loss.push({
          value_rem :"-",
          value_name : "rem_special"
        })
        }else{
           this.ac_remnants_loss.push({
          value_rem :this.actual_knit_special_rem,
          value_name : "rem_special"
        })
        }                 
          worksheet.getCell("O10").value =this.actual_knit_special_rem/100
          worksheet.getCell("O10").numFmt = '0.00%'; 
      

        this.total_actual_knit_woven = 0
         this.total_actual_knit_woven = Number(this.actual_knit_woven) + Number(this.actual_knit_woven_splice) 
         + Number(this.actual_knit_woven_cut)  + Number(this.actual_knit_woven_rem)

           worksheet.getCell("L11").value = this.total_actual_knit_woven/100
           worksheet.getCell("L11").numFmt = '0.00%'; 

           
        
    
        
        this.total_actual_knit_light = 0 
         this.total_actual_knit_light = Number(this.actual_knit_light) + Number(this.actual_knit_light_splice) 
         + Number(this.actual_knit_light_cut)  + Number(this.actual_knit_light_rem)

             worksheet.getCell("M11").value = this.total_actual_knit_light/100
             worksheet.getCell("M11").numFmt = '0.00%'; 
    

        this.total_actual_knit_fleece = 0
        this.total_actual_knit_fleece = Number(this.actual_knit_fleece) + Number(this.actual_knit_fleece_splice) 
         + Number(this.actual_knit_fleece_cut)  + Number(this.actual_knit_fleece_rem)

           worksheet.getCell("N11").value = this.total_actual_knit_fleece/100
           worksheet.getCell("N11").numFmt = '0.00%'; 
     

        this.total_actual_knit_special = 0
        this.total_actual_knit_special = Number(this.actual_knit_special) + Number(this.actual_knit_special_splice) 
         + Number(this.actual_knit_special_cut)  + Number(this.actual_knit_special_rem)

           worksheet.getCell("O11").value = this.total_actual_knit_special/100
           worksheet.getCell("O11").numFmt = '0.00%'; 
 

         worksheet.getCell("V1").value = ""; 
       worksheet.mergeCells("V1:V160");
       
        worksheet.getCell("V1").fill = {
          type: 'pattern',
          pattern:'solid',
          fgColor:{argb:'FF880000'},
          bgColor:{argb:'FF880000'}
          };
        
        

        

          worksheet.getCell("W3").value = "start"; 
      

           worksheet.getCell("W4").value = "25-Oct"; 
       

          this.total_act_woven_endloss_yd = 0 
          for(var axc = 0; axc<this.rowexport.length; axc++){
            if(this.rowexport[axc].f_type == "1")
            this.total_act_woven_endloss_yd = Number(this.total_act_woven_endloss_yd) + Number(this.rowexport[axc].endless_length_yd)
          }

          this.total_act_light_endloss_yd = 0 
          for(var axc = 0; axc<this.rowexport.length; axc++){
            if(this.rowexport[axc].f_type == "2")
            this.total_act_light_endloss_yd = Number(this.total_act_light_endloss_yd) + Number(this.rowexport[axc].endless_length_yd)
          }
          
          this.total_act_fleece_endloss_yd = 0 
          for(var axc = 0; axc<this.rowexport.length; axc++){
            if(this.rowexport[axc].f_type == "3")
            this.total_act_fleece_endloss_yd = Number(this.total_act_fleece_endloss_yd) + Number(this.rowexport[axc].endless_length_yd)
          }
          
           this.total_act_special_endloss_yd = 0 
          for(var axc = 0; axc<this.rowexport.length; axc++){
            if(this.rowexport[axc].f_type == "4")
            this.total_act_special_endloss_yd = Number(this.total_act_special_endloss_yd) + Number(this.rowexport[axc].endless_length_yd)
          }
       

        //end--------------------------------------------------------------------------------------------
            this.total_act_woven_spliceloss = 0
          for(var axc = 0; axc<this.rowexport.length; axc++){
            if(this.rowexport[axc].f_type == "1"){
            this.total_act_woven_spliceloss = Number(this.total_act_woven_spliceloss) + Number(this.rowexport[axc].spliceloss)
            }
          }
         
          this.total_act_light_spliceloss = 0
          for(var axc = 0; axc<this.rowexport.length; axc++){
            if(this.rowexport[axc].f_type == "2"){
            this.total_act_light_spliceloss = Number(this.total_act_light_spliceloss) + Number(this.rowexport[axc].spliceloss)
            }
          }

          this.total_act_fleece_spliceloss = 0
          for(var axc = 0; axc<this.rowexport.length; axc++){
            if(this.rowexport[axc].f_type == "3"){
            this.total_act_fleece_spliceloss = Number(this.total_act_fleece_spliceloss) + Number(this.rowexport[axc].spliceloss)
            }
          }

          this.total_act_special_spliceloss = 0
          for(var axc = 0; axc<this.rowexport.length; axc++){
            if(this.rowexport[axc].f_type == "4"){
            this.total_act_special_spliceloss = Number(this.total_act_special_spliceloss) + Number(this.rowexport[axc].spliceloss)
            }
          }

          //end--------------------------------------------------------------------------------------------
          this.total_act_woven_cut = 0
          for(var axc = 0; axc<this.rowexport.length; axc++){
            if(this.rowexport[axc].f_type == "1"){
            this.total_act_woven_cut = Number(this.total_act_woven_cut) + Number(this.rowexport[axc].totalcutout)
            }
          }

          this.total_act_light_cut = 0
          for(var axc = 0; axc<this.rowexport.length; axc++){
            if(this.rowexport[axc].f_type == "2"){
            this.total_act_light_cut = Number(this.total_act_light_cut) + Number(this.rowexport[axc].totalcutout)
            }
          }

           this.total_act_fleece_cut = 0
          for(var axc = 0; axc<this.rowexport.length; axc++){
            if(this.rowexport[axc].f_type == "3"){
            this.total_act_fleece_cut = Number(this.total_act_fleece_cut) + Number(this.rowexport[axc].totalcutout)
            }
          }

           this.total_act_special_cut = 0
          for(var axc = 0; axc<this.rowexport.length; axc++){
            if(this.rowexport[axc].f_type == "4"){
            this.total_act_special_cut = Number(this.total_act_special_cut) + Number(this.rowexport[axc].totalcutout)
            }
          }

           //end--------------------------------------------------------------------------------------------
           this.total_act_woven_rem = 0
          for(var axc = 0; axc<this.rowexport.length; axc++){
            if(this.rowexport[axc].f_type == "1"){
            this.total_act_woven_rem = Number(this.total_act_woven_rem) + Number(this.rowexport[axc].rement)
            }
          }

          this.total_act_light_rem = 0
          for(var axc = 0; axc<this.rowexport.length; axc++){
            if(this.rowexport[axc].f_type == "2"){
            this.total_act_light_rem = Number(this.total_act_light_rem) + Number(this.rowexport[axc].rement)
            }
          }

          this.total_act_fleece_rem = 0
          for(var axc = 0; axc<this.rowexport.length; axc++){
            if(this.rowexport[axc].f_type == "3"){
            this.total_act_fleece_rem = Number(this.total_act_fleece_rem) + Number(this.rowexport[axc].rement)
            }
          }

          this.total_act_special_rem = 0
          for(var axc = 0; axc<this.rowexport.length; axc++){
            if(this.rowexport[axc].f_type == "4"){
            this.total_act_special_rem = Number(this.total_act_special_rem) + Number(this.rowexport[axc].rement)
            }
          }
          //end--------------------------------------------------------------------------------------------

         
          this.sum_all_type_endloss_act = Number(this.total_act_woven_endloss_yd) + Number(this.total_act_light_endloss_yd)
          +Number(this.total_act_fleece_endloss_yd) + Number(this.total_act_special_endloss_yd)


           this.sum_all_type_woven_act = Number(this.total_act_woven_endloss_yd) + Number(this.total_act_woven_spliceloss)
          +Number(this.total_act_woven_cut) + Number(this.total_act_woven_rem)
          
            this.avg_act_endloss_yd  =  this.sum_all_type_endloss_act/this.all_type_summary_ttlmarker *100

            //end avg act endloss

        worksheet.getCell("P7").value = this.avg_act_endloss_yd/100
        worksheet.getCell("P7").numFmt = '0.00%'; 
  
      
          this.sum_all_type_splice_act = Number(this.total_act_woven_spliceloss) + Number(this.total_act_light_spliceloss)
          +Number(this.total_act_fleece_spliceloss) + Number(this.total_act_special_spliceloss)
          
          
          this.sum_all_type_light_act = Number(this.total_act_light_endloss_yd) + Number(this.total_act_light_spliceloss)
          +Number(this.total_act_light_cut) + Number(this.total_act_light_rem)
   
          
             this.avg_act_splice  =  this.sum_all_type_splice_act/this.all_type_summary_ttlmarker *100

            //end avg act splice

        worksheet.getCell("P8").value = this.avg_act_splice/100 
        worksheet.getCell("P8").numFmt = '0.00%'; 
  
         

          

          this.sum_all_type_cut_act = Number(this.total_act_woven_cut) + Number(this.total_act_light_cut)
          +Number(this.total_act_fleece_cut) + Number(this.total_act_special_cut)

         this.sum_all_type_fleece_act = Number(this.total_act_fleece_endloss_yd) + Number(this.total_act_fleece_spliceloss)
          +Number(this.total_act_fleece_cut) + Number(this.total_act_fleece_rem)
   
          
          this.avg_act_cut  =  this.sum_all_type_cut_act/this.all_type_summary_ttlmarker *100
    
        worksheet.getCell("P9").value = this.avg_act_cut/100
        worksheet.getCell("P9").numFmt = '0.00%';  
     

         
          this.sum_all_type_rem_act = Number(this.total_act_woven_rem) + Number(this.total_act_light_rem)
          +Number(this.total_act_fleece_rem) + Number(this.total_act_special_rem)

           this.sum_all_type_special_act = Number(this.total_act_special_endloss_yd) + Number(this.total_act_special_spliceloss)
          +Number(this.total_act_special_cut) + Number(this.total_act_special_rem)

          this.avg_act_rem  =  this.sum_all_type_rem_act/this.all_type_summary_ttlmarker *100
          
        worksheet.getCell("P10").value = this.avg_act_rem/100
        worksheet.getCell("P10").numFmt = '0.00%'; 
     
        this.total_act_audit = Number(this.sum_all_type_endloss_act) + Number(this.sum_all_type_splice_act) 
        + Number(this.sum_all_type_cut_act) + Number(this.sum_all_type_rem_act)
         
        this.total_act_spread_yd  = (Number(this.sum_all_type_endloss_act) + Number(this.sum_all_type_splice_act) 
        + Number(this.sum_all_type_cut_act) + Number(this.sum_all_type_rem_act))/this.all_type_summary_ttlmarker *100
      


          worksheet.getCell("P11").value = this.total_act_spread_yd/100
           worksheet.getCell("P11").numFmt = '0.00%'; 
   
           worksheet.getCell("P12").value = ""; 
    
           worksheet.getCell("P13").value = ""; 
     

      //Total  Spread  Loss Audit (Yds)
         worksheet.getCell("A12").value = "Total  Spread  Loss Audit (Yds)";
         worksheet.mergeCells("A12:K12");



        worksheet.getCell("L12").value = Number(this.sum_all_type_woven_act)
        worksheet.getCell("L12").numFmt = '0.00'; 
   

         worksheet.getCell("M12").value = Number(this.sum_all_type_light_act)
         worksheet.getCell("M12").numFmt = '0.00'; 


         worksheet.getCell("N12").value = Number(this.sum_all_type_fleece_act)
         worksheet.getCell("N12").numFmt = '0.00'; 
    

          worksheet.getCell("O12").value = Number(this.sum_all_type_special_act)
          worksheet.getCell("O12").numFmt = '0.00'; 
  

         worksheet.getCell("A13").value = "% of Fabric Mixture";
         worksheet.mergeCells("A13:K13");
    

       this.mixture_woven =  this.all_result_ttlmarker_woven/this.all_type_summary_ttlmarker *100
       worksheet.getCell("L13").value = this.mixture_woven/100
       worksheet.getCell("L13").numFmt = '0.00%'; 
       this.sum_all_4_ttl_marker.push({
             value:this.all_result_ttlmarker_woven,
             value_mix:this.mixture_woven,
             value_name:"Woven"
           })
     


        this.mixture_light =  this.all_result_ttlmarker_light/this.all_type_summary_ttlmarker *100
       worksheet.getCell("M13").value = this.mixture_light/100
       worksheet.getCell("M13").numFmt = '0.00%';
       this.sum_all_4_ttl_marker.push({
             value:this.all_result_ttlmarker_light,
             value_mix:this.mixture_light,
             value_name:"L&mid.wt"
           })
   
        
        this.mixture_fleece =  this.all_result_ttlmarker_fleece/this.all_type_summary_ttlmarker *100
       worksheet.getCell("N13").value = this.mixture_fleece/100
       worksheet.getCell("N13").numFmt = '0.00%'; 

       this.sum_all_4_ttl_marker.push({
             value:this.all_result_ttlmarker_fleece,
             value_mix:this.mixture_fleece,
             value_name:"Fleece"
           })
   
  
        this.mixture_special =  this.all_result_ttlmarker_special/this.all_type_summary_ttlmarker *100
       worksheet.getCell("O13").value = this.mixture_special/100
       worksheet.getCell("O13").numFmt = '0.00%';
       
       this.sum_all_4_ttl_marker.push({
             value:this.all_result_ttlmarker_special,
             value_mix:this.mixture_special,
             value_name:"Special"
           })


        //this.total_act_spread_yd

        worksheet.getCell("Q12").value = Number(this.total_act_audit)
        worksheet.getCell("Q12").numFmt = '0.00'; 
   

        
        const params2 = new FormData();
        params2.append("first_day",this.first_day)

        this.sum_day_week = (52*7)-1;
        console.log(this.sum_day_week)
        this.last_day = moment(this.first_day).add(this.sum_day_week,'days').format("YYYY/MM/DD");
        console.log(this.last_day)
      
        params2.append("last_day",this.last_day)
        params2.append("org",org)
        for (var pair of params2.entries()) {
            console.log(pair[0] + ", " + pair[1]);
      }

        await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data_all_weekly',
        data:params2,
      }).then((resp)=>{
       console.log(resp.data)
        this.allrowexport =[]
        
        function check_null(val){
                if(val == '' ||  val == null || val == "NaN" || val == "null" || val == undefined ){
                  return 0
                }else{
                  return val
                }
        }
       
        if(resp.data.data.length > 0){
             resp.data.data.forEach((e) => {
                this.allrowexport.push({
                  mu_date : e.MU_DATE,
                  gm_table:e.GM_TABLE,
                  so_no_doc:e.SO_NO_DOC,
                  customer:e.CUSTOMER_NAME,
                  gm_no:e.GM_NO,
                  gm_no_short :e.GM_NO.slice(2,6),
                  fabric_type:e.FABRIC_TYPE,
                  f_type:e.F_TYPE,
                  issue: e.ISSUE,
                  issue_roll :e.ROLL_NO,
                  avg_width:e.AVG_WIDTH,
                  gmm2: e.WEIGHT_M2,
                  gmyd: e.WEIGHT_YD,
                  width_inch : e.WIDTH_INCH,
                  length_ydb:e.YARD,
                  length_inch: e.LENGTH_INCH,
                  obs_width: e.AVG_WIDTH,
                  marker: e.MARKER,
                  roll:e.ROLL_NO,
                  ply: e.PLY,
                  endplug1: e.END_PLUG1,
                  endplug2: e.END_PLUG2,
                  cause: e.CAUSE_NAME,
                  code: e.CAUSE_CODE,
                  splicenoof: e.SPLICE,
                  spliceinch: e.SPLICE_INCH,
                  cutoutnoof: e.CUT_OUT,
                  cutoutweight: e.CUT_OUT_GRAM,
                  remnants: check_null(e.REMNANTS_WEIGHT),
                  widthplug1: e.PLUG1_GRAM,
                  widthplug2: e.PLUG2_GRAM,
                  avgglueside: e.GLUE_SIDE,
                  ttl_marker: parseFloat(e.TTL_MARKER_YD).toFixed(2),
                  yard_roll:e.YARD_ROLL,
                  length_yd:e.LENGTH_YD,
                  obswidth:e.OBS_WIDTH,
                  ttlobs_widthloss:e.OBS_WIDTH,
                  end1_2:(e.END1_2),
                  plug1_inch:e.PLUG1_INCH,
                  plug2_inch:e.PLUG2_INCH,
                  per_plug1:e.PER_PLUG1,
                  per_plug2:e.PER_PLUG2,
                  plug12:e.PLUG1_2_INCH,
                  endloss_yd: e.END_LOSS_YD,
                  widthloss:e.WITH_LOSS,
                  endless_length_yd:e.AVG_END_LOSS_YD,
                  avg_end:e.AVG_END_LOSS,
                  sectionlosslenth:check_null(e.SECTION_LOSS_YD),
                  sectionlossper:check_null(e.SECTION_LOSS_PER),
                  spliceloss:check_null(e.SPLICE_LOSS_YD),
                  splicelossper:parseFloat(check_null(e.SPLICE_LOSS_PER)).toFixed(2),
                  overlap:e.OVER_LAB_YD,
                  overlapper:parseFloat(e.OVER_LAB_PER).toFixed(2),
                  defect:parseFloat(e.DEFECT_YD).toFixed(2),
                  defectper:parseFloat(e.DEFECT_PER).toFixed(2),
                  totalcutout:parseFloat(check_null(e.TOTAL_CUTOUT_YD)).toFixed(2),
                  totalcutoutper:check_null(e.TOTAL_CUTOUT_PER),
                  rement:parseFloat(check_null(e.REMENT_LOSS_YD)).toFixed(2),
                  rement_per:parseFloat(check_null(e.REMENT_LOSS_PER)).toFixed(2),
                  totallenthloss:parseFloat(check_null(e.TOTAL_LENGTH_LOSS_YD)).toFixed(2),
                  totallenthlossper:parseFloat(check_null(e.TOTAL_LEN_LOSS_PER)).toFixed(2),
                  avg_end_code:e.AVG_END_CODE,
                  avg_end_cause:e.AVG_END_CAUSE,
                  per_splice_loss_code:e.SPLICE_LOSS_PER_CODE,
                  per_splice_loss_cause:e.SPLICE_LOSS_PER_CAUSE,
                  remnants_loss_code:e.REMNANTS_LOSS_CODE,
                  remnants_loss_cause:e.REMNANTS_LOSS_CAUSE,
                  total_cut_out_per_code:e.TOTAL_CUT_OUT_PER_CODE,
                  total_cut_out_per_cause:e.TOTAL_CUT_OUT_PER_CAUSE,
                  widthloss_code:e.WIDTH_LOSS_CODE,
                  widthloss_cause:e.WIDTH_LOSS_CAUSE,
                  balance: parseFloat(e.ISSUE-((e.TTL_MARKER_YD*e.WEIGHT_YD) /1000)).toFixed(2),
                })
               
       });
        }
        this.all_sum_cal = 0;
       
        for(var ax = 0; ax<this.allrowexport.length; ax++){
          this.all_sum_cal = Number(this.all_sum_cal) + Number(this.allrowexport[ax].ttl_marker)
        }
          worksheet.getCell("S6").value = Number(this.all_sum_cal);
          worksheet.getCell("S6").numFmt = "#,##0.00"

      
});
      
      var myNewDate = new Date(this.first_day);
        var myNewDate2 = new Date(myNewDate);
         var count = 0;
         this.ttl_marker_endless_woven = []
         this.ttl_marker_endless_light = []
         this.ttl_marker_endless_fleece = []
         this.ttl_marker_endless_special = []
        
        this.ttl_endloss_yd_woven = []
         this.ttl_endloss_yd_light = []
         this.ttl_endloss_yd_fleece = []
         this.ttl_endloss_yd_special = []
         
         this.ttl_splice_woven=[]
         this.ttl_splice_light=[]
         this.ttl_splice_fleece=[]
         this.ttl_splice_special=[]

         this.ttl_cut_woven=[]
         this.ttl_cut_light=[]
         this.ttl_cut_fleece=[]
         this.ttl_cut_special=[]

         this.ttl_rem_woven=[]
         this.ttl_rem_light=[]
         this.ttl_rem_fleece=[]
         this.ttl_rem_special=[]
         

         this.ttl_woven=[]
         this.ttl_light=[]
         this.ttl_fleece=[]
         this.ttl_special=[]

        this.end_target=[]
        this.splice_target=[]
        this.cut_target=[]
        this.rem_target=[]

        this.end_target_light=[]
      this.splice_target_light=[]
      this.cut_target_light=[]
      this.rem_target_light=[]

      this.end_target_fleece=[]
      this.splice_target_fleece=[]
      this.cut_target_fleece=[]
      this.rem_target_fleece=[]

      this.end_target_special=[]
      this.splice_target_special=[]
      this.cut_target_special=[]
      this.rem_target_special=[]

      this.total_ac_all_arr=[]

      var test = ["1","2","3"]
      test.push="4","5"
  

    
    
     for (var a1 = 1; a1 < 53; a1++) {
       this.ttl_marker_end_woven = 0
       this.ttl_marker_end_light = 0
       this.ttl_marker_end_fleece = 0
       this.ttl_marker_end_special= 0
       
       this.end_loss_yd_woven = 0
       this.end_loss_yd_light = 0
       this.end_loss_yd_fleece = 0
       this.end_loss_yd_special = 0
       
       this.splice_woven = 0
       this.splice_light = 0
       this.splice_fleece = 0
       this.splice_special = 0
      
       this.cut_woven = 0
       this.cut_light = 0
       this.cut_fleece = 0
       this.cut_special = 0

       this.rem_woven = 0
       this.rem_light = 0
       this.rem_fleece = 0
       this.rem_special = 0

       this.all_sum_widthloss = 0 
       this.all_sum_ttlmarker = 0
       

      
       
       
        
       if(count == 0 ){
          this.firstNewDate = moment(myNewDate).format("DD/MM");
          this.convert_axios1 = moment(myNewDate).format("YYYY-MM-DD");
          this.convert_week = moment(myNewDate).format("YYYY/MM/DD");
     

          this.next_date =  myNewDate.setDate(myNewDate.getDate() + 5);
          this.secNewDate = moment(myNewDate).format("DD/MM");
          this.convert_axios2 = moment(myNewDate).format("YYYY-MM-DD");
          this.convert_week2 = moment(myNewDate).format("YYYY/MM/DD");
  
          this.date_use_data.push({
            first_date:this.firstNewDate,
            sec_date:this.secNewDate,
            convert_axios1:this.convert_axios1,
            convert_axios2:this.convert_axios2,
            convert_week:this.convert_week,
            convert_week2:this.convert_week2,
            value:a1
          })
        
          count++
       }else{
          
         this.next_date = myNewDate.setDate(myNewDate.getDate() + 2);
          this.firstNewDate = moment(myNewDate).format("DD/MM");
          this.convert_axios1 = moment(myNewDate).format("YYYY-MM-DD");
          this.convert_week = moment(myNewDate).format("YYYY/MM/DD");
       

         
          this.next_date = myNewDate.setDate(myNewDate.getDate() + 5);
          this.secNewDate = moment(myNewDate).format("DD/MM");
          this.convert_axios2 = moment(myNewDate).format("YYYY-MM-DD");
          this.convert_week2 = moment(myNewDate).format("YYYY/MM/DD");
         
          this.date_use_data.push({
            first_date:this.firstNewDate,
            sec_date:this.secNewDate,
            convert_axios1:this.convert_axios1,
            convert_axios2:this.convert_axios2,
            convert_week:this.convert_week,
            convert_week2:this.convert_week2,
            value:a1
          })
          
          count++
       }
       
          const params3 = new FormData();
          params3.append("start",this.convert_axios1)
          params3.append("end",this.convert_axios2)
          let org = this.$q.localStorage.getItem("org");
          params3.append("org",org)
           await axios({
          method:'post',
          url:this.$api_url + "/mainso.php/export_data_5",
          data:params3
        }).then((resp)=>{
            
         if (resp.data.data.length > 0) {
           this.all_data_arr=[]
          resp.data.data.forEach((e) => {
            this.all_data_arr.push(
              {
              mu_date: e.MU_DATE,
              gm_table: e.GM_TABLE,
              so_no_doc: e.SO_NO_DOC,
              customer: e.CUSTOMER_NAME,
              gm_no: e.GM_NO,
              gm_no_short: e.GM_NO.slice(2, 6),
              fabric_type: e.FABRIC_TYPE,
              f_type: e.F_TYPE,
              issue: e.ISSUE,
              issue_roll: e.ROLL_NO,
              avg_width: e.AVG_WIDTH,
              gmm2: e.WEIGHT_M2,
              gmyd: e.WEIGHT_YD,
              width_inch: e.WIDTH_INCH,
              length_ydb: e.YARD,
              length_inch: e.LENGTH_INCH,
              obs_width: e.AVG_WIDTH,
              marker: e.MARKER,
              roll: e.ROLL_NO,
              ply: e.PLY,
              endplug1: e.END_PLUG1,
              endplug2: e.END_PLUG2,
              cause: e.CAUSE_NAME,
              code: e.CAUSE_CODE,
              splicenoof: e.SPLICE,
              spliceinch: e.SPLICE_INCH,
              cutoutnoof: e.CUT_OUT,
              cutoutweight: e.CUT_OUT_GRAM,
              remnants: e.REMNANTS_WEIGHT,
              widthplug1: e.PLUG1_GRAM,
              widthplug2: e.PLUG2_GRAM,
              avgglueside: e.GLUE_SIDE,
              ttl_marker: parseFloat(e.TTL_MARKER_YD).toFixed(2),
              yard_roll: e.YARD_ROLL,
              length_yd: e.LENGTH_YD,
              obswidth: e.OBS_WIDTH,
              ttlobs_widthloss: e.OBS_WIDTH,
              end1_2: e.END1_2,
              plug1_inch: e.PLUG1_INCH,
              plug2_inch: e.PLUG2_INCH,
              per_plug1: e.PER_PLUG1,
              per_plug2: e.PER_PLUG2,
              plug12: e.PLUG1_2_INCH,
              endloss_yd: e.END_LOSS_YD,
              widthloss: e.WITH_LOSS,
              endless_length_yd: e.AVG_END_LOSS_YD,
              avg_end: e.AVG_END_LOSS,
              sectionlosslenth: e.SECTION_LOSS_YD,
              sectionlossper: e.SECTION_LOSS_PER,
              spliceloss: e.SPLICE_LOSS_YD,
              splicelossper: e.SPLICE_LOSS_PER,
              overlap: e.OVER_LAB_YD,
              overlapper: e.OVER_LAB_PER,
              defect: e.DEFECT_YD,
              defectper: e.DEFECT_PER,
              totalcutout: e.TOTAL_CUTOUT_YD,
              totalcutoutper: e.TOTAL_CUTOUT_PER,
              rement: e.REMENT_LOSS_YD,
              rement_per: e.REMENT_LOSS_PER,
              totallenthloss: 
                e.TOTAL_LENGTH_LOSS_YD,
              totallenthlossper:
                e.TOTAL_LEN_LOSS_PER,
              avg_end_code: e.AVG_END_CODE,
              avg_end_cause: e.AVG_END_CAUSE,
              per_splice_loss_code: e.SPLICE_LOSS_PER_CODE,
              per_splice_loss_cause: e.SPLICE_LOSS_PER_CAUSE,
              remnants_loss_code: e.REMNANTS_LOSS_CODE,
              remnants_loss_cause: e.REMNANTS_LOSS_CAUSE,
              total_cut_out_per_code: e.TOTAL_CUT_OUT_PER_CODE,
              total_cut_out_per_cause: e.TOTAL_CUT_OUT_PER_CAUSE,
              widthloss_code: e.WIDTH_LOSS_CODE,
              widthloss_cause: e.WIDTH_LOSS_CAUSE,
              balance: parseFloat(
                e.ISSUE - (e.TTL_MARKER_YD * e.WEIGHT_YD) / 1000
              ).toFixed(2),
            })
         
           
            
              
          });
          this.sum_ttl_width = 0
            for(var ax = 0; ax<this.all_data_arr.length; ax++){
              this.all_sum_widthloss = 0
              this.all_sum_widthloss = this.all_data_arr[ax].ttl_marker * this.all_data_arr[ax].widthloss /100
              
              this.sum_ttl_width = Number(this.sum_ttl_width) + Number(this.all_sum_widthloss)

              this.all_sum_ttlmarker = Number(this.all_sum_ttlmarker) + Number(this.all_data_arr[ax].ttl_marker)
            
            }  
           
              this.sum_widthloss.push({
                    value:this.sum_ttl_width
                })

              this.sum_ttl_marker.push({
                    value: this.all_sum_ttlmarker
                })

      
             //start_ttl
            for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "1"){
              
            this.ttl_marker_end_woven = Number(this.ttl_marker_end_woven) + Number(this.all_data_arr[ax].ttl_marker)
            }
            }
          
            this.ttl_marker_endless_woven.push({
              end_ttl_marker_woven : this.ttl_marker_end_woven
            })

             for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "2"){
              
            this.ttl_marker_end_light = Number(this.ttl_marker_end_light) + Number(this.all_data_arr[ax].ttl_marker)
            }
            }
      
            this.ttl_marker_endless_light.push({
              end_ttl_marker_light : this.ttl_marker_end_light
            })

             for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "3"){
              
            this.ttl_marker_end_fleece = Number(this.ttl_marker_end_fleece) + Number(this.all_data_arr[ax].ttl_marker)
            }
            }
          
            this.ttl_marker_endless_fleece.push({
              end_ttl_marker_fleece : this.ttl_marker_end_fleece
            })

             for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "4"){
              
            this.ttl_marker_end_special = Number(this.ttl_marker_end_special) + Number(this.all_data_arr[ax].ttl_marker)
            }
            }
           
            this.ttl_marker_endless_special.push({
              end_ttl_marker_special : this.ttl_marker_end_special
            })
            //start_yd
             for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "1"){
              
            this.end_loss_yd_woven = Number(this.end_loss_yd_woven) + Number(this.all_data_arr[ax].endless_length_yd)
            }
            }
         
            this.ttl_endloss_yd_woven.push({
              end_loss_yd_woven : this.end_loss_yd_woven
            })

            for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "2"){
              
            this.end_loss_yd_light = Number(this.end_loss_yd_light) + Number(this.all_data_arr[ax].endless_length_yd)
            }
            }
         
            this.ttl_endloss_yd_light.push({
              end_loss_yd_light : this.end_loss_yd_light
            })

            for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "3"){
              
            this.end_loss_yd_fleece = Number(this.end_loss_yd_fleece) + Number(this.all_data_arr[ax].endless_length_yd)
            }
            }
          
            this.ttl_endloss_yd_fleece.push({
              end_loss_yd_fleece : this.end_loss_yd_fleece
            })

            for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "4"){
              
            this.end_loss_yd_special = Number(this.end_loss_yd_special) + Number(this.all_data_arr[ax].endless_length_yd)
            }
            }
            
            this.ttl_endloss_yd_special.push({
              end_loss_yd_special : this.end_loss_yd_special
            })
            //end_yd

            //start splice

            for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "1"){
              
            this.splice_woven = Number(this.splice_woven) + Number(this.all_data_arr[ax].spliceloss)
            }
            }
            
            this.ttl_splice_woven.push({
              splice_loss_woven : this.splice_woven
            })

             for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "2"){
              
            this.splice_light = Number(this.splice_light) + Number(this.all_data_arr[ax].spliceloss)
            }
            }
            
            this.ttl_splice_light.push({
              splice_loss_light : this.splice_light
            })

              for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "3"){
              
            this.splice_fleece = Number(this.splice_fleece) + Number(this.all_data_arr[ax].spliceloss)
            }
            }
            
            this.ttl_splice_fleece.push({
              splice_loss_fleece : this.splice_fleece
            })

            for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "4"){
              
            this.splice_special = Number(this.splice_special) + Number(this.all_data_arr[ax].spliceloss)
            }
            }
            
            this.ttl_splice_special.push({
              splice_loss_special : this.splice_special
            })

            //start cut


            for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "1"){
              
            this.cut_woven = Number(this.cut_woven) + Number(this.all_data_arr[ax].totalcutout)
            }
            }
            
            this.ttl_cut_woven.push({
              cut_loss_woven : this.cut_woven
            })

             for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "2"){
              
            this.cut_light = Number(this.cut_light) + Number(this.all_data_arr[ax].totalcutout)
            }
            }
            
            this.ttl_cut_light.push({
              cut_loss_light : this.cut_light
            })

              for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "3"){
              
            this.cut_fleece = Number(this.cut_fleece) + Number(this.all_data_arr[ax].totalcutout)
            }
            }
            
            this.ttl_cut_fleece.push({
              cut_loss_fleece : this.cut_fleece
            })

            for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "4"){
              
            this.cut_special = Number(this.cut_special) + Number(this.all_data_arr[ax].totalcutout)
            }
            }
            
            this.ttl_cut_special.push({
              cut_loss_special : this.cut_special
            })

            //end cut


            //start rem
            
            for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "1"){
            
            this.rem_woven = Number(this.rem_woven) + Number(this.all_data_arr[ax].rement)
            }
            }
           
            this.ttl_rem_woven.push({
              rem_loss_woven : this.rem_woven
            })
           
             for(var ax = 0; ax<this.all_data_arr.length; ax++){
               
            if(this.all_data_arr[ax].f_type == "2"){
         
            this.rem_light = Number(this.rem_light) + Number(this.all_data_arr[ax].rement)
        
            }
            }
            
            this.ttl_rem_light.push({
              rem_loss_light : this.rem_light
            })
           
          
              for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "3"){
              
            this.rem_fleece = Number(this.rem_fleece) + Number(this.all_data_arr[ax].rement)
            }
            }
           
            this.ttl_rem_fleece.push({
              rem_loss_fleece : this.rem_fleece
            })
        
            for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
            if(this.all_data_arr[ax].f_type == "4"){
              
            this.rem_special = Number(this.rem_special) + Number(this.all_data_arr[ax].rement)
            }
            }
            
            this.ttl_rem_special.push({
              rem_loss_special : this.rem_special
            })
          
        
            
         }else{
            this.all_data_arr.push({
                 mu_date: "",
                gm_table: "",
                so_no_doc: "",
                customer: "",
                gm_no: "",
                gm_no_short: "",
                fabric_type: "",
                f_type: "",
                issue: "",
                issue_roll: "",
                avg_width: "",
                gmm2: "",
                gmyd: "",
                width_inch: "",
                length_ydb: "",
                length_inch: "",
                obs_width: "",
                marker: "",
                roll: "",
                ply: "",
                endplug1: "",
                endplug2: "",
                cause: "",
                code: "",
                splicenoof: "",
                spliceinch: "",
                cutoutnoof: "",
                cutoutweight: "",
                remnants: "",
                widthplug1: "",
                widthplug2: "",
                avgglueside: "",
                ttl_marker: "",
                yard_roll: "",
                length_yd: "",
                obswidth: "",
                ttlobs_widthloss: "",
                end1_2: "",
                plug1_inch: "",
                plug2_inch: "",
                per_plug1: "",
                per_plug2: "",
                plug12: "",
                endloss_yd: "",
                widthloss: "",
                endless_length_yd: "",
                avg_end: "",
                sectionlosslenth: "",
                sectionlossper: "",
                spliceloss: "",
                splicelossper: "",
                overlap: "",
                overlapper: "",
                defect: "",
                defectper: "",
                totalcutout: "",
                totalcutoutper: "",
                rement: "",
                rement_per: "",
                totallenthloss: "",
                totallenthlossper: "",
                avg_end_code: "",
                avg_end_cause: "",
                per_splice_loss_code: "",
                per_splice_loss_cause: "",
                remnants_loss_code: "",
                remnants_loss_cause: "",
                total_cut_out_per_code: "",
                total_cut_out_per_cause: "",
                widthloss_code: "",
                widthloss_cause: "",

            })
          
        this.ttl_marker_endless_woven.push({
              end_ttl_marker_woven : 0
        })
        this.ttl_marker_endless_light.push({
              end_ttl_marker_light : 0
        })
         this.ttl_marker_endless_fleece.push({
              end_ttl_marker_fleece : 0
        })
         this.ttl_marker_endless_special.push({
              end_ttl_marker_special : 0
        })

        this.ttl_endloss_yd_woven.push({
              end_loss_yd_woven : 0
        })

        this.ttl_endloss_yd_light.push({
              end_loss_yd_light : 0
        })

        this.ttl_endloss_yd_fleece.push({
              end_loss_yd_fleece : 0
        })

        this.ttl_endloss_yd_special.push({
              end_loss_yd_special : 0
        })
        this.ttl_splice_woven.push({
              splice_loss_woven : 0
        })
        this.ttl_splice_light.push({
              splice_loss_light : 0
        })
        this.ttl_splice_fleece.push({
              splice_loss_fleece : 0
        })
        this.ttl_splice_special.push({
              splice_loss_special :0
        })
        this.ttl_cut_woven.push({
              cut_loss_woven : 0
            })
        this.ttl_cut_light.push({
              cut_loss_light : 0
            })
        this.ttl_cut_fleece.push({
              cut_loss_fleece : 0
            })
        this.ttl_cut_special.push({
              cut_loss_special : 0
            })
          this.ttl_rem_woven.push({
              rem_loss_woven : 0
            })
        this.ttl_rem_light.push({
              rem_loss_light : 0
            })
        this.ttl_rem_fleece.push({
              rem_loss_fleece : 0
            })
        this.ttl_rem_special.push({
              rem_loss_special : 0
            })
        this.sum_widthloss.push({
            value:0
            })
        this.sum_ttl_marker.push({
            value:0
            })
         }
        })
         
        
     }  
 

       function change_date(val){
         let day = val.substring(0, 2);
         let month = val.substring(3, 5);
         if(month == "1"){
           month = "JAN"
         }else if(month == "2"){
           month = "FEB"
         }else if(month == "3"){
           month = "MAR"
         }else if(month == "4"){
           month = "APR"
         }else if(month == "5"){
           month = "MAY"
         }else if(month == "6"){
           month = "JUN"
         }else if(month == "7"){
           month = "JUL"
         }else if(month == "8"){
           month = "AUG"
         }else if(month == "9"){
           month = "SEP"
         }else if(month == "10"){
           month = "OCT"
         }else if(month == "11"){
           month = "NOV"
         }else{
           month = "DEC"
         }
         let change_day = day + "-" + month
         return change_day
       }

     
        worksheet.getCell("X1").value = "Week Calculator";
        worksheet.getCell("X1").alignment = {
          
          
        };
        worksheet.getCell("X1").border = {
         
          
          
          
        };
        
        for(var ax=0; ax<this.arr_colume.length; ax++){
           worksheet.getCell(this.arr_colume[ax].name_col+"2").value = "Week"+[ax+1];
     

        }

    
        for(var ax=0; ax<this.arr_colume.length; ax++){
        worksheet.getCell(this.arr_colume[ax].name_col+"3").value = this.date_use_data[ax].first_date + "-" +this.date_use_data[ax].sec_date;
   
        }

        for(var ax=0; ax<this.arr_colume.length; ax++){
        worksheet.getCell(this.arr_colume[ax].name_col+"4").value =  change_date(this.date_use_data[ax].sec_date);
  
        }


        //ส่วนที่2

        for(var ax=0; ax<this.arr_colume.length; ax++){
           worksheet.getCell(this.arr_colume[ax].name_col+"31").value = [ax+1];
       

        }


          worksheet.getCell("X33").value = 1;
       


        for(var ax=0; ax<this.arr_colume.length; ax++){
           worksheet.getCell(this.arr_colume[ax].name_col+"34").value = "Week"+[ax+1];
       

        }

        for(var ax=0; ax<this.arr_colume.length; ax++){
        worksheet.getCell(this.arr_colume[ax].name_col+"35").value = this.date_use_data[ax].first_date + "-" +this.date_use_data[ax].sec_date;
     
        }

             for(var ax=0; ax<this.arr_colume.length; ax++){
          
          
        worksheet.getCell(this.arr_colume[ax].name_col+"36").value = Number(this.ttl_marker_endless_woven[ax].end_ttl_marker_woven)
        worksheet.getCell(this.arr_colume[ax].name_col+"36").numFmt = '#,##0.00'; 
      worksheet.getCell(this.arr_colume[ax].name_col+"36").fill = {
          type: 'pattern',
          pattern:'solid',
          fgColor:{argb:'FF33FF66'},
          bgColor:{argb:'FF33FF66'}
          };
     
        }
      
        for(var ax=0; ax<this.arr_colume.length; ax++){
       
          this.total_spread_loss_woven_target.push({
          value:0.38
        })
        worksheet.getCell(this.arr_colume[ax].name_col+"42").value = 0.38/100;
        worksheet.getCell(this.arr_colume[ax].name_col+"42").numFmt = '0.00%';
     
        }

        


       worksheet.getCell("W36").value = "Total  Marker  yd";
    


         worksheet.getCell("W37").value = "End loss";
      
        
         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.end_loss_woven_per = 0
           this.end_loss_woven_per = this.ttl_endloss_yd_woven[ax].end_loss_yd_woven/this.ttl_marker_endless_woven[ax].end_ttl_marker_woven*100
           
           if(this.end_loss_woven_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"37").value = this.end_loss_woven_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"37").numFmt = '0.00%';
  
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"37").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"37").numFmt = '0.00%';
       
           }
         }

         worksheet.getCell("W38").value = "Splice Loss";
         

        for(var ax=0; ax<this.arr_colume.length; ax++){
           this.splice_woven_per = 0
           this.splice_woven_per = this.ttl_splice_woven[ax].splice_loss_woven/this.ttl_marker_endless_woven[ax].end_ttl_marker_woven*100
           if(this.splice_woven_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"38").value = this.splice_woven_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"38").numFmt = '0.00%';
       
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"38").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"38").numFmt = '0.00%';
      
           }
         }

              worksheet.getCell("W39").value = "Cut Out Loss";
     

         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.cut_woven_per = 0
           this.cut_woven_per = this.ttl_cut_woven[ax].cut_loss_woven/this.ttl_marker_endless_woven[ax].end_ttl_marker_woven*100
           if(this.cut_woven_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"39").value = this.cut_woven_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"39").numFmt = '0.00%';
       
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"39").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"39").numFmt = '0.00%';
       
           }
         }



              worksheet.getCell("W40").value = "Remnants Loss";
       

        for(var ax=0; ax<this.arr_colume.length; ax++){
           this.rem_woven_per = 0
           this.rem_woven_per = this.ttl_rem_woven[ax].rem_loss_woven/this.ttl_marker_endless_woven[ax].end_ttl_marker_woven*100
           if(this.rem_woven_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"40").value = this.rem_woven_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"40").numFmt = '0.00%';
  
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"40").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"40").numFmt = '0.00%';
     
           }
         }


              worksheet.getCell("W41").value = "% Total  Spread  Loss ";
      

        worksheet.getCell("W42").value = "% Total  Spread  Loss  Target";
   


        worksheet.getCell("W43").value = "Total  Spread  Loss  (Yds)";
  

        worksheet.getCell("W45").value = "Total Marker ( yd )";
   

        for(var ax=0; ax<this.arr_colume.length; ax++){
          
          
        worksheet.getCell(this.arr_colume[ax].name_col+"45").value = Number(this.ttl_marker_endless_woven[ax].end_ttl_marker_woven);
        worksheet.getCell(this.arr_colume[ax].name_col+"45").numFmt = '#,##0.00';
      
      worksheet.getCell(this.arr_colume[ax].name_col+"45").fill = {
          type: 'pattern',
          pattern:'solid',
          fgColor:{argb:'FF33FF66'},
          bgColor:{argb:'FF33FF66'}
          };
       
       
        }


        worksheet.getCell("W46").value = "Total Act. End Loss (yd)";
      
       

        
         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.end_loss_woven_per = 0
           this.end_loss_woven_per = this.ttl_endloss_yd_woven[ax].end_loss_yd_woven
           this.end_target_woven.push({
             value_target_woven:this.end_loss_woven_per
           })
           if(this.end_loss_woven_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"46").value = Number(this.end_loss_woven_per)
          worksheet.getCell(this.arr_colume[ax].name_col+"46").numFmt = '0.00';
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"46").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"46").numFmt = '0.00';
     
           }
         }
        


          worksheet.getCell("W47").value = "Total Act. Splice Loss (yd)";
     

         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.splice_woven_per = 0
           this.splice_woven_per = this.ttl_splice_woven[ax].splice_loss_woven
            this.splice_target_woven.push({
             value_target_woven:this.splice_woven_per
           })
           
           if(this.splice_woven_per > 0 && this.splice_woven_per != undefined && isNaN(this.splice_light_per)==false){
          worksheet.getCell(this.arr_colume[ax].name_col+"47").value = Number(this.splice_light_per)
          worksheet.getCell(this.arr_colume[ax].name_col+"47").numFmt = '0.00';
      
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"47").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"47").numFmt = '0.00';
        
           }
         }

          worksheet.getCell("W48").value = "Total Act. Cut Out Loss (yd)";
     

         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.cut_woven_per = 0
           this.cut_woven_per = this.ttl_cut_woven[ax].cut_loss_woven
               this.cut_target_woven.push({
             value_target_woven:this.cut_woven_per
           })
           if(this.cut_woven_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"48").value = Number(this.cut_woven_per);
          worksheet.getCell(this.arr_colume[ax].name_col+"48").numFmt = '0.00';
       
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"48").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"48").numFmt = '0.00';
      
           }
         }

          worksheet.getCell("W49").value = "Total Act. Remnants Loss (yd)";
       

         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.rem_woven_per = 0
           this.rem_woven_per = this.ttl_rem_woven[ax].rem_loss_woven
                 this.rem_target_woven.push({
             value_target_woven:this.rem_woven_per
           })
          
           if(this.rem_light_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"49").value = Number(this.rem_light_per)
          worksheet.getCell(this.arr_colume[ax].name_col+"49").numFmt = '0.00';
     
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"49").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"46").numFmt = '0.00';
      
           }
         }

          worksheet.getCell("W50").value = "Total Act.  Spread Loss  (yd)";
      

         
        for(var ax=0; ax<this.arr_colume.length; ax++){
           this.all_ttl_act = 0
           this.ttl_spread_loss_woven = 0
           this.all_ttl_act = Number(this.ttl_endloss_yd_woven[ax].end_loss_yd_woven) + Number(this.ttl_splice_woven[ax].splice_loss_woven) + Number(this.ttl_cut_woven[ax].cut_loss_woven) 
           +Number(this.ttl_rem_woven[ax].rem_loss_woven)
           
           this.ttl_woven.push({
             value_all_woven : this.all_ttl_act
           })
           if(this.all_ttl_act > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"50").value = Number(this.all_ttl_act);
          worksheet.getCell(this.arr_colume[ax].name_col+"50").numFmt = '0.00';
        
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"50").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"50").numFmt = '0.00';
       
           }


             if(this.all_ttl_act > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"43").value = Number(this.all_ttl_act);
          worksheet.getCell(this.arr_colume[ax].name_col+"43").numFmt = '0.00';
        
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"43").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"43").numFmt = '0.00';
        
           }


        this.ttl_spread_loss_woven = this.all_ttl_act / this.ttl_marker_endless_woven[ax].end_ttl_marker_woven *100
  
        if(isNaN(this.ttl_spread_loss_woven)==true){
        this.ttl_spread_loss_woven = 0
        }
        if(this.ttl_spread_loss_woven > 0){
         this.total_spread_loss_woven_data.push({
           value:this.ttl_spread_loss_woven,
           value_ax:ax
         })
        }
        
        if(this.ttl_spread_loss_woven > 0){
        worksheet.getCell(this.arr_colume[ax].name_col+"41").value = this.ttl_spread_loss_woven/100
        worksheet.getCell(this.arr_colume[ax].name_col+"41").numFmt = '0.00%';
       
        }else{
         worksheet.getCell(this.arr_colume[ax].name_col+"41").value = 0.00;
         worksheet.getCell(this.arr_colume[ax].name_col+"41").numFmt = '0.00%';
        
        }
        }

        //end row woven


//strat row light
          worksheet.getCell("X53").value = 2;
     

          for(var ax=0; ax<this.arr_colume.length; ax++){
           worksheet.getCell(this.arr_colume[ax].name_col+"54").value = "Week"+[ax+1];
    

        }


        for(var ax=0; ax<this.arr_colume.length; ax++){
        worksheet.getCell(this.arr_colume[ax].name_col+"55").value = this.date_use_data[ax].first_date + "-" +this.date_use_data[ax].sec_date;
      
        }


        worksheet.getCell("W56").value = "Total  Marker  yd";
      
  
         
         for(var ax=0; ax<this.arr_colume.length; ax++){
          
          
        worksheet.getCell(this.arr_colume[ax].name_col+"56").value = Number(this.ttl_marker_endless_light[ax].end_ttl_marker_light)
        worksheet.getCell(this.arr_colume[ax].name_col+"56").numFmt = '#,##0.00';
         
      worksheet.getCell(this.arr_colume[ax].name_col+"56").fill = {
          type: 'pattern',
          pattern:'solid',
          fgColor:{argb:'FF33FF66'},
          bgColor:{argb:'FF33FF66'}
          };
       
        }


         worksheet.getCell("W57").value = "End loss";
   

        
         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.end_loss_light_per = 0
           this.end_loss_light_per = this.ttl_endloss_yd_light[ax].end_loss_yd_light/this.ttl_marker_endless_light[ax].end_ttl_marker_light*100
           if(this.end_loss_light_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"57").value = this.end_loss_light_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"57").numFmt = '0.00%';
      
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"57").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"57").numFmt = '0.00%';
       
           }
         }
          

         worksheet.getCell("W58").value = "Splice Loss";
   


        for(var ax=0; ax<this.arr_colume.length; ax++){
           this.splice_light_per = 0
           this.splice_light_per = this.ttl_splice_light[ax].splice_loss_light/this.ttl_marker_endless_light[ax].end_ttl_marker_light*100
           if(this.splice_light_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"58").value = this.splice_light_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"58").numFmt = '0.00%';
      
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"58").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"58").numFmt = '0.00%';
        
           }
         }

              worksheet.getCell("W59").value = "Cut Out Loss";
        
        

        for(var ax=0; ax<this.arr_colume.length; ax++){
           this.cut_light_per = 0
           this.cut_light_per = this.ttl_cut_light[ax].cut_loss_light/this.ttl_marker_endless_light[ax].end_ttl_marker_light*100
           if(this.cut_light_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"59").value = this.cut_light_per/100;
          worksheet.getCell(this.arr_colume[ax].name_col+"59").numFmt = '0.00%';
   
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"59").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"59").numFmt = '0.00%';
       
           }
         }

              worksheet.getCell("W60").value = "Remnants Loss";
    


        for(var ax=0; ax<this.arr_colume.length; ax++){
           this.rem_light_per = 0
           this.rem_light_per = this.ttl_rem_light[ax].rem_loss_light/this.ttl_marker_endless_light[ax].end_ttl_marker_light*100
           if(this.rem_light_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"60").value = this.rem_light_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"60").numFmt = '0.00%';
    
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"60").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"60").numFmt = '0.00%';

      
           }
         }


              worksheet.getCell("W61").value = "% Total  Spread  Loss ";
  

   


        worksheet.getCell("W62").value = "% Total  Spread  Loss  Target";


        //this.Light_total

        
        for(var ax=0; ax<this.arr_colume.length; ax++){
        
          this.total_spread_loss_light_target.push({
          value:this.Light_total
        })
          worksheet.getCell(this.arr_colume[ax].name_col+"62").value = this.Light_total/100
          worksheet.getCell(this.arr_colume[ax].name_col+"62").numFmt = '0.00%';
        }


        worksheet.getCell("W63").value = "Total  Spread  Loss  (Yds)";



        worksheet.getCell("W65").value = "Total Marker ( yd )";


        for(var ax=0; ax<this.arr_colume.length; ax++){
          
          
        worksheet.getCell(this.arr_colume[ax].name_col+"65").value = Number(this.ttl_marker_endless_light[ax].end_ttl_marker_light)
        worksheet.getCell(this.arr_colume[ax].name_col+"65").numFmt = '#,##0.00';
 
      worksheet.getCell(this.arr_colume[ax].name_col+"65").fill = {
          type: 'pattern',
          pattern:'solid',
          fgColor:{argb:'FF33FF66'},
          bgColor:{argb:'FF33FF66'}
          };
 
        }


        worksheet.getCell("W66").value = "Total Act. End Loss (yd)";
   

        
         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.end_loss_light_per = 0
           this.end_loss_light_per = this.ttl_endloss_yd_light[ax].end_loss_yd_light
             this.end_target_light.push({
             value_target_light:this.end_loss_light_per
           })
           if(this.end_loss_light_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"66").value = Number(this.end_loss_light_per)
          worksheet.getCell(this.arr_colume[ax].name_col+"66").numFmt = '0.00';
 
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"66").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"66").numFmt = '0.00';
   
           }
         }
        


          worksheet.getCell("W67").value = "Total Act. Splice Loss (yd)";
  

         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.splice_light_per = 0
           this.splice_light_per = this.ttl_splice_light[ax].splice_loss_light
             this.splice_target_light.push({
             value_target_light:this.splice_light_per
           })
           if(this.splice_light_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"67").value = Number(this.splice_light_per)
          worksheet.getCell(this.arr_colume[ax].name_col+"67").numFmt = '0.00';
     
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"67").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"67").numFmt = '0.00';
   
           }
         }

          worksheet.getCell("W68").value = "Total Act. Cut Out Loss (yd)";
  


         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.cut_light_per = 0
           this.cut_light_per = this.ttl_cut_light[ax].cut_loss_light
                this.cut_target_light.push({
             value_target_light:this.cut_light_per
           })
           if(this.cut_light_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"68").value = Number(this.cut_light_per)
          worksheet.getCell(this.arr_colume[ax].name_col+"68").numFmt = '0.00';
    
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"68").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"68").numFmt = '0.00';
     
           }
         }

          worksheet.getCell("W69").value = "Total Act. Remnants Loss (yd)";
   

         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.rem_light_per = 0
           this.rem_light_per = this.ttl_rem_light[ax].rem_loss_light
                  this.rem_target_light.push({
             value_target_light:this.rem_light_per
           })
           if(this.rem_light_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"69").value = Number(this.rem_light_per)
          worksheet.getCell(this.arr_colume[ax].name_col+"69").numFmt = '0.00';
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"69").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"69").numFmt = '0.00';
       
           }
         }

          worksheet.getCell("W70").value = "Total Act.  Spread Loss  (yd)";
      

         
        for(var ax=0; ax<this.arr_colume.length; ax++){
           this.all_ttl_act = 0
           this.ttl_spread_loss_light = 0
           this.all_ttl_act = Number(this.ttl_endloss_yd_light[ax].end_loss_yd_light) + Number(this.ttl_splice_light[ax].splice_loss_light) + Number(this.ttl_cut_light[ax].cut_loss_light) 
           +Number(this.ttl_rem_light[ax].rem_loss_light)

            this.ttl_light.push({
             value_all_light : this.all_ttl_act
           }) 
          
           if(this.all_ttl_act > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"70").value = Number(this.all_ttl_act)
          worksheet.getCell(this.arr_colume[ax].name_col+"70").numFmt = '0.00';
     
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"70").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"70").numFmt = '0.00';
     
           }


             if(this.all_ttl_act > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"63").value = Number(this.all_ttl_act)
          worksheet.getCell(this.arr_colume[ax].name_col+"63").numFmt = '0.00';
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"63").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"63").numFmt = '0.00';
   
           }


        this.ttl_spread_loss_light = this.all_ttl_act / this.ttl_marker_endless_light[ax].end_ttl_marker_light *100
        if(isNaN(this.ttl_spread_loss_light)==true){
        this.ttl_spread_loss_light = 0
        }
        this.total_spread_loss_light_data.push({
          value:this.ttl_spread_loss_light
        })
        if(this.ttl_spread_loss_light > 0){
        worksheet.getCell(this.arr_colume[ax].name_col+"61").value = this.ttl_spread_loss_light/100
        worksheet.getCell(this.arr_colume[ax].name_col+"61").numFmt = '0.00%';
      
        }else{
         worksheet.getCell(this.arr_colume[ax].name_col+"61").value = 0.00/100;
         worksheet.getCell(this.arr_colume[ax].name_col+"61").numFmt = '0.00%';
    
        }
        }
        

        //end_all_row 2

    //start row fleece
          worksheet.getCell("X74").value = 3;
      

          for(var ax=0; ax<this.arr_colume.length; ax++){
           worksheet.getCell(this.arr_colume[ax].name_col+"75").value = "Week"+[ax+1];
     

        }


        for(var ax=0; ax<this.arr_colume.length; ax++){
        worksheet.getCell(this.arr_colume[ax].name_col+"76").value = this.date_use_data[ax].first_date + "-" +this.date_use_data[ax].sec_date;
 
        }


        worksheet.getCell("W77").value = "Total  Marker  yd";
 
  
         
         for(var ax=0; ax<this.arr_colume.length; ax++){
          
          
        worksheet.getCell(this.arr_colume[ax].name_col+"77").value = Number(this.ttl_marker_endless_fleece[ax].end_ttl_marker_fleece)
        worksheet.getCell(this.arr_colume[ax].name_col+"77").numFmt = '#,##0.00';
         
      worksheet.getCell(this.arr_colume[ax].name_col+"77").fill = {
          type: 'pattern',
          pattern:'solid',
          fgColor:{argb:'FF33FF66'},
          bgColor:{argb:'FF33FF66'}
          };
    
        }


         worksheet.getCell("W78").value = "End loss";
     

        
         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.end_loss_fleece_per = 0
           this.end_loss_fleece_per = this.ttl_endloss_yd_fleece[ax].end_loss_yd_fleece/this.ttl_marker_endless_fleece[ax].end_ttl_marker_fleece*100
           if(this.end_loss_fleece_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"78").value = this.end_loss_fleece_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"78").numFmt = '0.00%';
   
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"78").value = 0.00/100;
               worksheet.getCell(this.arr_colume[ax].name_col+"78").numFmt = '0.00';
     
           }
         }
          

         worksheet.getCell("W79").value = "Splice Loss";
      


        for(var ax=0; ax<this.arr_colume.length; ax++){
           this.splice_fleece_per = 0
           this.splice_fleece_per = this.ttl_splice_fleece[ax].splice_loss_fleece/this.ttl_marker_endless_fleece[ax].end_ttl_marker_fleece*100
           if(this.splice_fleece_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"79").value = this.splice_fleece_per/100;
          worksheet.getCell(this.arr_colume[ax].name_col+"79").numFmt = '0.00%';
      
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"79").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"79").numFmt = '0.00%';

           }
         }

              worksheet.getCell("W80").value = "Cut Out Loss";
     
        

        for(var ax=0; ax<this.arr_colume.length; ax++){
           this.cut_fleece_per = 0
           this.cut_fleece_per = this.ttl_cut_fleece[ax].cut_loss_fleece/this.ttl_marker_endless_fleece[ax].end_ttl_marker_fleece*100
           if(this.cut_fleece_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"80").value = this.cut_fleece_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"80").numFmt = '0.00%';
     
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"80").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"80").numFmt = '0.00%';
 
           }
         }



              worksheet.getCell("W81").value = "Remnants Loss";



        for(var ax=0; ax<this.arr_colume.length; ax++){
           this.rem_fleece_per = 0
           this.rem_fleece_per = this.ttl_rem_fleece[ax].rem_loss_fleece/this.ttl_marker_endless_fleece[ax].end_ttl_marker_fleece*100
        
           if(this.rem_fleece_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"81").value = this.rem_fleece_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"81").numFmt = '0.00%';
 
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"81").value = 0.00/100;
               worksheet.getCell(this.arr_colume[ax].name_col+"81").numFmt = '0.00%';
   
           }
         }


              worksheet.getCell("W82").value = "% Total  Spread  Loss ";
    
   


        worksheet.getCell("W83").value = "% Total  Spread  Loss  Target";
    

        //this.Fleece_total

        
        for(var ax=0; ax<this.arr_colume.length; ax++){
            this.total_spread_loss_fleece_target.push({
            value:this.Fleece_total
          })
          
          worksheet.getCell(this.arr_colume[ax].name_col+"83").value = this.Fleece_total/100;
          worksheet.getCell(this.arr_colume[ax].name_col+"83").numFmt = '0.00%';
        }

     
        worksheet.getCell("W84").value = "Total  Spread  Loss  (Yds)";
    


        worksheet.getCell("W86").value = "Total Marker ( yd )";
  

        for(var ax=0; ax<this.arr_colume.length; ax++){
          
          
        worksheet.getCell(this.arr_colume[ax].name_col+"86").value = Number(this.ttl_marker_endless_fleece[ax].end_ttl_marker_fleece);
        worksheet.getCell(this.arr_colume[ax].name_col+"86").numFmt = '#,##0.00';
      
      worksheet.getCell(this.arr_colume[ax].name_col+"86").fill = {
          type: 'pattern',
          pattern:'solid',
          fgColor:{argb:'FF33FF66'},
          bgColor:{argb:'FF33FF66'}
          };
  
        }


        worksheet.getCell("W87").value = "Total Act. End Loss (yd)";
   

        
         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.end_loss_fleece_per = 0
           this.end_loss_fleece_per = this.ttl_endloss_yd_fleece[ax].end_loss_yd_fleece
              this.end_target_fleece.push({
             value_target_fleece:this.end_loss_fleece_per
           })
           if(this.end_loss_fleece_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"87").value = Number(this.end_loss_fleece_per)
          worksheet.getCell(this.arr_colume[ax].name_col+"87").numFmt = '0.00';
     
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"87").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"87").numFmt = '0.00';
   
           }
         }
        


          worksheet.getCell("W88").value = "Total Act. Splice Loss (yd)";
  

         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.splice_fleece_per = 0
           this.splice_fleece_per = this.ttl_splice_fleece[ax].splice_loss_fleece
             this.splice_target_fleece.push({
             value_target_fleece:this.splice_fleece_per
           })
           if(this.splice_fleece_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"88").value = Number(this.splice_fleece_per)
          worksheet.getCell(this.arr_colume[ax].name_col+"88").numFmt = '0.00';
    
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"88").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"88").numFmt = '0.00';

           }
         }

          worksheet.getCell("W89").value = "Total Act. Cut Out Loss (yd)";
    

        
         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.cut_fleece_per = 0
           this.cut_fleece_per = this.ttl_cut_fleece[ax].cut_loss_fleece
               this.cut_target_fleece.push({
             value_target_fleece:this.cut_fleece_per
           })
           if(this.cut_fleece_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"89").value = Number(this.cut_fleece_per)
          worksheet.getCell(this.arr_colume[ax].name_col+"89").numFmt = '0.00';
   
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"89").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"89").numFmt = '0.00';
  
           }
         }

          worksheet.getCell("W90").value = "Total Act. Remnants Loss (yd)";
  

         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.rem_fleece_per = 0
           this.rem_fleece_per = this.ttl_rem_fleece[ax].rem_loss_fleece
                   this.rem_target_fleece.push({
             value_target_fleece:this.rem_fleece_per
           })
           if(this.rem_fleece_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"90").value = Number(this.rem_fleece_per);
          worksheet.getCell(this.arr_colume[ax].name_col+"90").numFmt = '0.00';
     
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"90").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"90").numFmt = '0.00';
     
           }
         }

          worksheet.getCell("W91").value = "Total Act.  Spread Loss  (yd)";
        

         
        for(var ax=0; ax<this.arr_colume.length; ax++){
           this.all_ttl_act = 0
           this.ttl_spread_loss_fleece = 0
           this.all_ttl_act = Number(this.ttl_endloss_yd_fleece[ax].end_loss_yd_fleece) + Number(this.ttl_splice_fleece[ax].splice_loss_fleece) + Number(this.ttl_cut_fleece[ax].cut_loss_fleece) 
           +Number(this.ttl_rem_fleece[ax].rem_loss_fleece)

            this.ttl_fleece.push({
             value_all_fleece : this.all_ttl_act
           })
           
           if(this.all_ttl_act > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"91").value = Number(this.all_ttl_act)
          worksheet.getCell(this.arr_colume[ax].name_col+"91").numFmt = '0.00';
   
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"91").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"91").numFmt = '0.00';
      
           }


             if(this.all_ttl_act > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"84").value = Number(this.all_ttl_act)
          worksheet.getCell(this.arr_colume[ax].name_col+"84").numFmt = '0.00';
       
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"84").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"84").numFmt = '0.00';
      
           }

  
        this.ttl_spread_loss_fleece = this.all_ttl_act / this.ttl_marker_endless_fleece[ax].end_ttl_marker_fleece *100
         if(isNaN(this.ttl_spread_loss_fleece )==true){
        this.ttl_spread_loss_fleece  = 0
        }
             this.total_spread_loss_fleece_data.push({
             value:this.ttl_spread_loss_fleece  
           })
        if(this.ttl_spread_loss_fleece > 0){
         
        
        worksheet.getCell(this.arr_colume[ax].name_col+"82").value = this.ttl_spread_loss_fleece/100;
        worksheet.getCell(this.arr_colume[ax].name_col+"82").numFmt = '0.00%';
     
        }else{
         worksheet.getCell(this.arr_colume[ax].name_col+"82").value = 0.00;
         worksheet.getCell(this.arr_colume[ax].name_col+"82").numFmt = '0.00%';
 
        } 
       }
      //end 3


      //start row special
       worksheet.getCell("X94").value = 4;
     

          for(var ax=0; ax<this.arr_colume.length; ax++){
           worksheet.getCell(this.arr_colume[ax].name_col+"95").value = "Week"+[ax+1];


        }


        for(var ax=0; ax<this.arr_colume.length; ax++){
        worksheet.getCell(this.arr_colume[ax].name_col+"96").value = this.date_use_data[ax].first_date + "-" +this.date_use_data[ax].sec_date;

        }


        worksheet.getCell("W97").value = "Total  Marker  yd";
   
  
         
         for(var ax=0; ax<this.arr_colume.length; ax++){
          
          
        worksheet.getCell(this.arr_colume[ax].name_col+"97").value = Number(this.ttl_marker_endless_special[ax].end_ttl_marker_special)
        worksheet.getCell(this.arr_colume[ax].name_col+"97").numFmt = '#,##0.00';
      
      worksheet.getCell(this.arr_colume[ax].name_col+"97").fill = {
          type: 'pattern',
          pattern:'solid',
          fgColor:{argb:'FF33FF66'},
          bgColor:{argb:'FF33FF66'}
          };
     
        }


         worksheet.getCell("W98").value = "End loss";
     

        
         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.end_loss_special_per = 0
           this.end_loss_special_per = this.ttl_endloss_yd_special[ax].end_loss_yd_special/this.ttl_marker_endless_special[ax].end_ttl_marker_special*100
           if(this.end_loss_special_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"98").value = this.end_loss_special_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"98").numFmt = '0.00%';
       
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"98").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"98").numFmt = '0.00%';
  
           }
         }
          

         worksheet.getCell("W99").value = "Splice Loss";
      


        for(var ax=0; ax<this.arr_colume.length; ax++){
           this.splice_special_per = 0
           this.splice_special_per = this.ttl_splice_special[ax].splice_loss_special/this.ttl_marker_endless_special[ax].end_ttl_marker_special*100
           if(this.splice_special_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"99").value = this.splice_special_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"99").numFmt = '0.00%';
       
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"99").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"99").numFmt = '0.00%';
   
           }
         }

              worksheet.getCell("W100").value = "Cut Out Loss";
       
        
        

        for(var ax=0; ax<this.arr_colume.length; ax++){
           this.cut_special_per = 0
           this.cut_special_per = this.ttl_cut_special[ax].cut_loss_special/this.ttl_marker_endless_special[ax].end_ttl_marker_special*100
           if(this.cut_special_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"100").value = this.cut_special_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"100").numFmt = '0.00%';

          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"100").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"100").numFmt = '0.00%';
     
           }
         }



              worksheet.getCell("W101").value = "Remnants Loss";
      
      


        for(var ax=0; ax<this.arr_colume.length; ax++){
           this.rem_special_per = 0
           this.rem_special_per = this.ttl_rem_special[ax].rem_loss_special/this.ttl_marker_endless_special[ax].end_ttl_marker_special*100
           if(this.rem_special_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"101").value = this.rem_special_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"101").numFmt = '0.00%';
   
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"101").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"101").numFmt = '0.00%';
     
    
           }
         }


              worksheet.getCell("W102").value = "% Total  Spread  Loss ";
  

   


        worksheet.getCell("W103").value = "% Total  Spread  Loss  Target";
  

        //this.Special_total

        
        for(var ax=0; ax<this.arr_colume.length; ax++){
          this.total_spread_loss_special_target.push({
          value:this.Special_total
        })
          worksheet.getCell(this.arr_colume[ax].name_col+"103").value = this.Special_total/100
          worksheet.getCell(this.arr_colume[ax].name_col+"103").numFmt = '0.00%';
     

        }


        worksheet.getCell("W104").value = "Total  Spread  Loss  (Yds)";
      

        worksheet.getCell("W106").value = "Total Marker ( yd )";
      

        for(var ax=0; ax<this.arr_colume.length; ax++){
          
          
        worksheet.getCell(this.arr_colume[ax].name_col+"106").value = Number(this.ttl_marker_endless_special[ax].end_ttl_marker_special)
        worksheet.getCell(this.arr_colume[ax].name_col+"106").numFmt = '#,##0.00';
      
      worksheet.getCell(this.arr_colume[ax].name_col+"106").fill = {
          type: 'pattern',
          pattern:'solid',
          fgColor:{argb:'FF33FF66'},
          bgColor:{argb:'FF33FF66'}
          };
 
        }


        worksheet.getCell("W107").value = "Total Act. End Loss (yd)";
   

        
         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.end_loss_special_per = 0
           this.end_loss_special_per = this.ttl_endloss_yd_special[ax].end_loss_yd_special
             this.end_target_special.push({
             value_target_special:this.end_loss_special_per
           })
           if(this.end_loss_special_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"107").value = Number(this.end_loss_special_per)
          worksheet.getCell(this.arr_colume[ax].name_col+"107").numFmt = '0.00';
   
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"107").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"106").numFmt = '0.00';
     
           }
         }
        


          worksheet.getCell("W108").value = "Total Act. Splice Loss (yd)";
   

         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.splice_special_per = 0
           this.splice_special_per = this.ttl_splice_special[ax].splice_loss_special
             this.splice_target_special.push({
             value_target_special:this.splice_special_per
           })
           if(this.splice_special_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"108").value = Number(this.splice_special_per)
          worksheet.getCell(this.arr_colume[ax].name_col+"108").numFmt = '0.00';
     
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"108").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"108").numFmt = '0.00';
    
           }
         }

          worksheet.getCell("W109").value = "Total Act. Cut Out Loss (yd)";
    


         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.cut_special_per = 0
           this.cut_special_per = this.ttl_cut_special[ax].cut_loss_special
                 this.cut_target_special.push({
             value_target_special:this.cut_special_per
           })
           if(this.cut_special_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"109").value = Number(this.cut_special_per)
          worksheet.getCell(this.arr_colume[ax].name_col+"109").numFmt = '0.00';
      
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"109").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"109").numFmt = '0.00';
  
           }
         }

          worksheet.getCell("W110").value = "Total Act. Remnants Loss (yd)";
        

         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.rem_special_per = 0
           this.rem_special_per = this.ttl_rem_special[ax].rem_loss_special
                      this.rem_target_special.push({
             value_target_special:this.rem_special_per
           })
           if(this.rem_special_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"110").value = Number(this.rem_special_per)
          worksheet.getCell(this.arr_colume[ax].name_col+"110").numFmt = '0.00';
     
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"110").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"90").numFmt = '0.00';
      
           }
         }

          worksheet.getCell("W111").value = "Total Act.  Spread Loss  (yd)";
     
        

         
        for(var ax=0; ax<this.arr_colume.length; ax++){
           this.all_ttl_act = 0
           this.ttl_spread_loss_special = 0
           this.all_ttl_act = Number(this.ttl_endloss_yd_special[ax].end_loss_yd_special) + Number(this.ttl_splice_special[ax].splice_loss_special) + Number(this.ttl_cut_special[ax].cut_loss_special) 
           +Number(this.ttl_rem_special[ax].rem_loss_special)
              this.ttl_special.push({
             value_all_special : this.all_ttl_act
           })
          
           if(this.all_ttl_act > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"111").value = Number(this.all_ttl_act)
          worksheet.getCell(this.arr_colume[ax].name_col+"111").numFmt = '0.00';
    
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"111").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"111").numFmt = '0.00';
      
           }


             if(this.all_ttl_act > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"104").value = Number(this.all_ttl_act);
          worksheet.getCell(this.arr_colume[ax].name_col+"104").numFmt = '0.00';
      
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"104").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"104").numFmt = '0.00';
  
           }

        
        this.ttl_spread_loss_special = this.all_ttl_act / this.ttl_marker_endless_special[ax].end_ttl_marker_special *100
            if(isNaN(this.ttl_spread_loss_special )==true){
        this.ttl_spread_loss_special  = 0
        }
        this.total_spread_loss_special_data.push({
          value:this.ttl_spread_loss_special
        })
        if(this.ttl_spread_loss_special > 0){
        worksheet.getCell(this.arr_colume[ax].name_col+"102").value = this.ttl_spread_loss_special/100
        worksheet.getCell(this.arr_colume[ax].name_col+"102").numFmt = '0.00%';
   
        }else{
         worksheet.getCell(this.arr_colume[ax].name_col+"102").value = 0.00;
         worksheet.getCell(this.arr_colume[ax].name_col+"102").numFmt = '0.00%';
 
        } 
        }
        //end row special

        //start Total

          worksheet.getCell("X114").value = "1-4";
    

          for(var ax=0; ax<this.arr_colume.length; ax++){
           worksheet.getCell(this.arr_colume[ax].name_col+"115").value = "Week"+[ax+1];
 

        }


        for(var ax=0; ax<this.arr_colume.length; ax++){
        worksheet.getCell(this.arr_colume[ax].name_col+"116").value = this.date_use_data[ax].first_date + "-" +this.date_use_data[ax].sec_date;

        }


        worksheet.getCell("W117").value = "Total  Marker  yd";
  
  
         
         for(var ax=0; ax<this.arr_colume.length; ax++){
          this.total_ttl_marker1_4 = 0
          this.total_ttl_marker1_4 = Number(this.ttl_marker_endless_woven[ax].end_ttl_marker_woven) + Number(this.ttl_marker_endless_light[ax].end_ttl_marker_light)
          +Number(this.ttl_marker_endless_fleece[ax].end_ttl_marker_fleece) + Number(this.ttl_marker_endless_special[ax].end_ttl_marker_special)
           this.total_target.push({
             all_ttl : this.total_ttl_marker1_4
           })
        worksheet.getCell(this.arr_colume[ax].name_col+"117").value = Number(this.total_ttl_marker1_4)
        worksheet.getCell(this.arr_colume[ax].name_col+"117").numFmt = '#,##0.00';

       
  
      worksheet.getCell(this.arr_colume[ax].name_col+"117").fill = {
          type: 'pattern',
          pattern:'solid',
          fgColor:{argb:'FF33FF66'},
          bgColor:{argb:'FF33FF66'}
          };
 
        }

        worksheet.getCell("W130").value = "Total  Marker  yd";
  
  
         
         for(var ax=0; ax<this.arr_colume.length; ax++){
          this.total_ttl_marker1_4 = 0
          this.total_ttl_marker1_4 = Number(this.ttl_marker_endless_woven[ax].end_ttl_marker_woven) + Number(this.ttl_marker_endless_light[ax].end_ttl_marker_light)
          +Number(this.ttl_marker_endless_fleece[ax].end_ttl_marker_fleece) + Number(this.ttl_marker_endless_special[ax].end_ttl_marker_special)
           this.total_target.push({
             all_ttl : this.total_ttl_marker1_4
           })
        worksheet.getCell(this.arr_colume[ax].name_col+"130").value = Number(this.total_ttl_marker1_4)
        worksheet.getCell(this.arr_colume[ax].name_col+"130").numFmt = '#,##0.00';
  
      worksheet.getCell(this.arr_colume[ax].name_col+"130").fill = {
          type: 'pattern',
          pattern:'solid',
          fgColor:{argb:'FF33FF66'},
          bgColor:{argb:'FF33FF66'}
          };
 
        }


  
   
        this.total_ttl_marker1_4
        
      

         worksheet.getCell("W119").value = "Splice Loss";
   

        
      

              worksheet.getCell("W120").value = "Cut Out Loss";

        

       



              worksheet.getCell("W121").value = "Remnants Loss";
 

      
        

              worksheet.getCell("W122").value = "% Total  Spread  Loss ";

   
      

        worksheet.getCell("W123").value = "% Total  Spread  Loss  Target";
 

        //this.Special_total
      
        
        for(var ax=0; ax<this.arr_colume.length; ax++){
        
          this.total_spread_target = 0
          this.total_spread_target = (Number(this.total_audit[0].value_total * this.ttl_marker_endless_woven[ax].end_ttl_marker_woven)  + Number(this.total_audit[1].value_total * this.ttl_marker_endless_light[ax].end_ttl_marker_light)
          +Number(this.total_audit[2].value_total * this.ttl_marker_endless_fleece[ax].end_ttl_marker_fleece) + Number(this.total_audit[3].value_total *  this.ttl_marker_endless_special[ax].end_ttl_marker_special)) /this.total_target[ax].all_ttl  
         
          if(this.total_spread_target > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"123").value = this.total_spread_target/100
          worksheet.getCell(this.arr_colume[ax].name_col+"123").numFmt = '0.00%';

          }else{
             worksheet.getCell(this.arr_colume[ax].name_col+"123").value = 0.00;
             worksheet.getCell(this.arr_colume[ax].name_col+"123").numFmt = '0.00%';

          }
      }

    /*     worksheet.getCell("W125").value = "Total Marker ( yd )";
        worksheet.getCell("W125").font = {
        
        
        
        
        
      };
        worksheet.getCell("W125").alignment = {
          
          
        };
        worksheet.getCell("W125").border = {
         
          
          
          
        }; */

        


        worksheet.getCell("W125").value = "Total Act. End Loss (yd)";


        
         for(var ax=0; ax<this.arr_colume.length; ax++){

             this.total_ac_end_loss_1_4 = 0
         
      
        this.total_ac_end_loss_1_4 = Number(this.end_target_woven[ax].value_target_woven) + Number(this.end_target_light[ax].value_target_light) 
        + Number(this.end_target_fleece[ax].value_target_fleece) +Number(this.end_target_special[ax].value_target_special)
          
         this.total_ac_end_arr.push({
           ac_end : this.total_ac_end_loss_1_4
         })
          
           if(this.total_ac_end_loss_1_4 > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"125").value = Number(this.total_ac_end_loss_1_4);
          worksheet.getCell(this.arr_colume[ax].name_col+"125").numFmt = '#,##0.00';
          worksheet.getCell(this.arr_colume[ax].name_col+"125").fill = {
          type: 'pattern',
          pattern:'solid',
          fgColor:{argb:'FF33FF66'},
          bgColor:{argb:'FF33FF66'}
          };
 
        

          }else{
          worksheet.getCell(this.arr_colume[ax].name_col+"125").value = 0.00;
          worksheet.getCell(this.arr_colume[ax].name_col+"125").numFmt = '#,##0.00';
          worksheet.getCell(this.arr_colume[ax].name_col+"125").fill = {
          type: 'pattern',
          pattern:'solid',
          fgColor:{argb:'FF33FF66'},
          bgColor:{argb:'FF33FF66'}
          };
 
        
           }
         }

            for(var ax=0; ax<this.arr_colume.length; ax++){

          this.end_loss_special_per = 0
           this.end_loss_special_per = this.total_ac_end_arr[ax].ac_end/this.total_target[ax].all_ttl*100
          this.arr_end_loss_1_4.push({
            value:this.end_loss_special_per
          })
           if(this.end_loss_special_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"118").value = this.end_loss_special_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"118").numFmt = '0.00%';


          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"118").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"118").numFmt = '0.00%';

           }
        
           }

          worksheet.getCell("W118").value = "End Loss";
           for(var ax=0; ax<this.arr_colume.length; ax++){

          this.end_loss_special_per = 0
           this.end_loss_special_per = this.total_ac_end_arr[ax].ac_end/this.total_target[ax].all_ttl*100
          this.arr_end_loss_1_4.push({
            value:this.end_loss_special_per
          })
           if(this.end_loss_special_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"118").value = this.end_loss_special_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"118").numFmt = '0.00%';


          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"118").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"118").numFmt = '0.00%';

           }
        
           }

          worksheet.getCell("W126").value = "Total Act. Splice Loss (yd)";


         for(var ax=0; ax<this.arr_colume.length; ax++){
             this.total_ac_splice_loss_1_4 = 0
         
      
        this.total_ac_splice_loss_1_4 = Number(this.splice_target_woven[ax].value_target_woven) + Number(this.splice_target_light[ax].value_target_light) 
        + Number(this.splice_target_fleece[ax].value_target_fleece) +Number(this.splice_target_special[ax].value_target_special)
          
          this.total_ac_splice_arr.push({
           ac_splice : this.total_ac_splice_loss_1_4
         })
           if(this.total_ac_splice_loss_1_4 > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"126").value = Number(this.total_ac_splice_loss_1_4)
          worksheet.getCell(this.arr_colume[ax].name_col+"126").numFmt = '0.00';

          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"126").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"126").numFmt = '0.00';

           }
         }

          worksheet.getCell("W127").value = "Total Act. Cut Out Loss (yd)";

            for(var ax=0; ax<this.arr_colume.length; ax++){
           this.splice_special_per = 0
          
   
           this.splice_special_per = this.total_ac_splice_arr[ax].ac_splice/this.total_target[ax].all_ttl*100
           if(this.splice_special_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"119").value = this.splice_special_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"119").numFmt = '0.00%';

          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"119").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"119").numFmt = '0.00%';

           }
         }


        
         for(var ax=0; ax<this.arr_colume.length; ax++){
            this.total_ac_cut_loss_1_4 = 0
         
  
      
        this.total_ac_cut_loss_1_4 = Number(this.cut_target_woven[ax].value_target_woven) + Number(this.cut_target_light[ax].value_target_light) 
        + Number(this.cut_target_fleece[ax].value_target_fleece) +Number(this.cut_target_special[ax].value_target_special)
           this.total_ac_cut_arr.push({
           ac_cut : this.total_ac_cut_loss_1_4
         })
           if(this.cut_special_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"127").value = Number(this.total_ac_cut_loss_1_4);
          worksheet.getCell(this.arr_colume[ax].name_col+"127").numFmt = '0.00';

          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"127").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"127").numFmt = '0.00';

           }
         }

         for(var ax=0; ax<this.arr_colume.length; ax++){
           this.cut_special_per = 0
           this.cut_special_per = this.total_ac_cut_arr[ax].ac_cut/this.total_target[ax].all_ttl*100
           if(this.cut_special_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"120").value = this.cut_special_per/100;
          worksheet.getCell(this.arr_colume[ax].name_col+"120").numFmt = '0.00%';

          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"120").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"120").numFmt = '0.00%';

           }
         }



          worksheet.getCell("W128").value = "Total Act. Remnants Loss (yd)";

        
         for(var ax=0; ax<this.arr_colume.length; ax++){
            
         this.total_ac_rem_loss_1_4 = 0
     

        this.total_ac_rem_loss_1_4 = Number(this.rem_target_woven[ax].value_target_woven) + Number(this.rem_target_light[ax].value_target_light) 
        + Number(this.rem_target_fleece[ax].value_target_fleece) +Number(this.rem_target_special[ax].value_target_special)

      
           this.total_ac_rem_arr.push({
           ac_rem : this.total_ac_rem_loss_1_4
         })
           if(this.total_ac_rem_loss_1_4 > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"128").value = Number(this.total_ac_rem_loss_1_4)
          worksheet.getCell(this.arr_colume[ax].name_col+"128").numFmt = '0.00';
 
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"128").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"126").numFmt = '0.00';

           }
         }

          for(var ax=0; ax<this.arr_colume.length; ax++){
           this.rem_special_per = 0
           this.rem_special_per = this.total_ac_rem_arr[ax].ac_rem/this.total_target[ax].all_ttl*100
           if(this.rem_special_per > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"121").value = this.rem_special_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"121").numFmt = '0.00%';

          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"121").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"121").numFmt = '0.00%';

           }
         }
      
    

          worksheet.getCell("W129").value = "Total Act.  Spread Loss  (yd)";
  

         
        for(var ax=0; ax<this.arr_colume.length; ax++){
        this.total_act_spread_yd_1_4 = 0
        this.total_act_spread_yd_1_4 = Number(this.total_ac_end_arr[ax].ac_end) + Number(this.total_ac_splice_arr[ax].ac_splice)+Number(this.total_ac_cut_arr[ax].ac_cut)
        +Number(this.total_ac_rem_arr[ax].ac_rem)
            this.total_ac_all_arr.push({
              value_all_ac : this.total_act_spread_yd_1_4
            })
           
           if(this.total_act_spread_yd_1_4 > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"129").value = Number(this.total_act_spread_yd_1_4)
          worksheet.getCell(this.arr_colume[ax].name_col+"129").numFmt = '0.00';

          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"129").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"129").numFmt = '0.00';

           }

          



           

  /* this.total_ac_all_arr.push({
              value_all_ac : this.total_act_spread_yd_1_4
            })

            this.total_target.push({
             all_ttl : this.total_ttl_marker1_4
           }) */
        
     
        this.ttl_spread_loss_special = this.total_ac_all_arr[ax].value_all_ac / this.total_target[ax].all_ttl *100
        if(this.ttl_spread_loss_special > 0){
        worksheet.getCell(this.arr_colume[ax].name_col+"122").value = this.ttl_spread_loss_special/100
        worksheet.getCell(this.arr_colume[ax].name_col+"122").numFmt = '0.00%';

        }else{
         worksheet.getCell(this.arr_colume[ax].name_col+"122").value = 0.00;
         worksheet.getCell(this.arr_colume[ax].name_col+"122").numFmt = '0.00%';
 
        } 
        }


        worksheet.getCell("W131").value = "Fabric Type";
 

        
        worksheet.getCell("X131").value = "% Target Phase5";
      


        worksheet.getCell("Z131").value = "Phase 6 Target(Current)";
 

         worksheet.getCell("W137").value = "Total Act.  Spread Loss  (yd)";


        for(var ax=0; ax<this.arr_colume.length; ax++){
        this.total_act_spread_yd_1_4 = 0
        this.total_act_spread_yd_1_4 = Number(this.total_ac_end_arr[ax].ac_end) + Number(this.total_ac_splice_arr[ax].ac_splice)+Number(this.total_ac_cut_arr[ax].ac_cut)
        +Number(this.total_ac_rem_arr[ax].ac_rem)
          
           if(this.total_act_spread_yd_1_4 > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"137").value = Number(this.total_act_spread_yd_1_4);
          worksheet.getCell(this.arr_colume[ax].name_col+"137").numFmt = '0.00';
  
          }else{
               worksheet.getCell(this.arr_colume[ax].name_col+"137").value = 0.00;
               worksheet.getCell(this.arr_colume[ax].name_col+"137").numFmt = '0.00';

           }
        }

         worksheet.getCell("W138").value = "% Total  Spread  Loss  Target";
    

        for(var ax=0; ax<this.arr_colume.length; ax++){
      
           this.total_spread_target = 0
          this.total_spread_target = (Number(this.total_audit[0].value_total * this.ttl_marker_endless_woven[ax].end_ttl_marker_woven)  + Number(this.total_audit[1].value_total * this.ttl_marker_endless_light[ax].end_ttl_marker_light)
          +Number(this.total_audit[2].value_total * this.ttl_marker_endless_fleece[ax].end_ttl_marker_fleece) + Number(this.total_audit[3].value_total *  this.ttl_marker_endless_special[ax].end_ttl_marker_special)) /this.total_target[ax].all_ttl  
         
          if(this.total_spread_target > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"138").value = this.total_spread_target/100;
          worksheet.getCell(this.arr_colume[ax].name_col+"138").numFmt = '0.00%';
    
          }else{
             worksheet.getCell(this.arr_colume[ax].name_col+"138").value = 0.00;
             worksheet.getCell(this.arr_colume[ax].name_col+"138").numFmt = '0.00%';
 
          }
        }
        
        this.cell = 132

     
        for(var ax= 0; ax<this.phase_type.length; ax++){
          this.cell = 132 + ax

        worksheet.getCell("W"+this.cell).value = ax+1;
   

        worksheet.getCell("X"+this.cell).value = this.phase_type[ax].per/100
         worksheet.getCell("X"+this.cell).numFmt = '0.00%';
   
         worksheet.getCell("Z"+this.cell).value = Number(this.total_audit[ax].value_total)
         worksheet.getCell("Z"+this.cell).numFmt = '0.00';

        }

         worksheet.getCell("X141").value = "Target Weight 1-4";

        
          for(var ax=0; ax<this.arr_colume.length; ax++){
           worksheet.getCell(this.arr_colume[ax].name_col+"142").value = "Week"+[ax+1];
 

        }


        for(var ax=0; ax<this.arr_colume.length; ax++){
        worksheet.getCell(this.arr_colume[ax].name_col+"143").value = this.date_use_data[ax].first_date + "-" +this.date_use_data[ax].sec_date;

        }

         worksheet.getCell("W144").value = "End Loss Target";
  

        for(var ax=0; ax<this.arr_colume.length; ax++){
     
          this.total_end_target = 0
          this.total_end_target = (Number(this.end_loss_woven * this.ttl_marker_endless_woven[ax].end_ttl_marker_woven)  + Number(this.end_loss_light * this.ttl_marker_endless_light[ax].end_ttl_marker_light)
          +Number(this.end_loss_fleece * this.ttl_marker_endless_fleece[ax].end_ttl_marker_fleece) + Number(this.end_loss_special *  this.ttl_marker_endless_special[ax].end_ttl_marker_special)) /this.total_target[ax].all_ttl  
          this.last_total_spread_end.push({
            value_last_spread : this.total_end_target
          })
          if(this.total_end_target > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"144").value = this.total_end_target/100
          worksheet.getCell(this.arr_colume[ax].name_col+"144").numFmt = '0.00%'; 

          }else{
             worksheet.getCell(this.arr_colume[ax].name_col+"144").value = 0.00;
             worksheet.getCell(this.arr_colume[ax].name_col+"144").numFmt = '0.00%'; 
 
          }
      }

         

         worksheet.getCell("W145").value = "Splice Loss Target";
   

        for(var ax=0; ax<this.arr_colume.length; ax++){
     
          this.total_splice_target = 0
          this.total_splice_target = (Number(this.splice_loss_woven * this.ttl_marker_endless_woven[ax].end_ttl_marker_woven)  + Number(this.splice_loss_light * this.ttl_marker_endless_light[ax].end_ttl_marker_light)
          +Number(this.splice_loss_fleece * this.ttl_marker_endless_fleece[ax].end_ttl_marker_fleece) + Number(this.splice_loss_special *  this.ttl_marker_endless_special[ax].end_ttl_marker_special)) /this.total_target[ax].all_ttl  
         this.last_total_spread_splice.push({
            value_last_spread : this.total_splice_target
          })
          if(this.total_splice_target > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"145").value = this.total_splice_target/100
          worksheet.getCell(this.arr_colume[ax].name_col+"145").numFmt = '0.00%'; 
  
          }else{
             worksheet.getCell(this.arr_colume[ax].name_col+"145").value = 0.00;
             worksheet.getCell(this.arr_colume[ax].name_col+"145").numFmt = '0.00%'; 
    
          }
      }


        worksheet.getCell("W146").value = "Cut out Loss Target";
     

         for(var ax=0; ax<this.arr_colume.length; ax++){
     
          this.total_cut_target = 0
          this.total_cut_target = (Number(this.cut_loss_woven * this.ttl_marker_endless_woven[ax].end_ttl_marker_woven)  + Number(this.cut_loss_light * this.ttl_marker_endless_light[ax].end_ttl_marker_light)
          +Number(this.cut_loss_fleece * this.ttl_marker_endless_fleece[ax].end_ttl_marker_fleece) + Number(this.cut_loss_special *  this.ttl_marker_endless_special[ax].end_ttl_marker_special)) /this.total_target[ax].all_ttl  
         
         this.last_total_spread_cut.push({
            value_last_spread : this.total_cut_target
          })
          if(this.total_cut_target > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"146").value = this.total_cut_target/100;
          worksheet.getCell(this.arr_colume[ax].name_col+"146").numFmt = '0.00%'; 
 
          }else{
             worksheet.getCell(this.arr_colume[ax].name_col+"146").value = 0.00;
             worksheet.getCell(this.arr_colume[ax].name_col+"146").numFmt = '0.00%'; 

          }
      }


        worksheet.getCell("W147").value = "Remnants Loss Target";
       
         for(var ax=0; ax<this.arr_colume.length; ax++){
     
          this.total_rem_target = 0
          
//แก้ remก่อน
          this.total_rem_target = (Number(this.rem_loss_woven * this.ttl_marker_endless_woven[ax].end_ttl_marker_woven)  + Number(this.rem_loss_light * this.ttl_marker_endless_light[ax].end_ttl_marker_light)
          +Number(this.rem_loss_fleece * this.ttl_marker_endless_fleece[ax].end_ttl_marker_fleece) + Number(this.rem_loss_special *  this.ttl_marker_endless_special[ax].end_ttl_marker_special))/this.total_target[ax].all_ttl 
         
         this.last_total_spread_rem.push({
            value_last_spread : this.total_rem_target
          })
          if(this.total_rem_target > 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"147").value = this.total_rem_target/100
          worksheet.getCell(this.arr_colume[ax].name_col+"147").numFmt = '0.00%'; 
       
   
          }else{
             worksheet.getCell(this.arr_colume[ax].name_col+"147").value = 0.00;
             worksheet.getCell(this.arr_colume[ax].name_col+"147").numFmt = '0.00%'; 
        
        
        
          }
      }
         
       
        
        worksheet.getCell("W148").value = "% Total  Spread  Loss  Target";
        
      
       

         for(var ax=0; ax<this.arr_colume.length; ax++){
          this.all_last_spread = 0
          this.all_last_spread = Number(this.last_total_spread_end[ax].value_last_spread) + Number(this.last_total_spread_splice[ax].value_last_spread)
          +Number(this.last_total_spread_cut[ax].value_last_spread) + Number(this.last_total_spread_rem[ax].value_last_spread)
        
          if(this.all_last_spread > 0){

          worksheet.getCell(this.arr_colume[ax].name_col+"148").value = this.all_last_spread/100
          worksheet.getCell(this.arr_colume[ax].name_col+"148").numFmt = '0.00%'; 

          }else{
            worksheet.getCell(this.arr_colume[ax].name_col+"148").value = 0.00
             worksheet.getCell(this.arr_colume[ax].name_col+"148").numFmt = '0.00%'; 

          }
      }
 
      
       
        
      
     
      

         
       this.find_max_ttl_marker = 0
      this.find_max_ttl_marker_name = ""
      this.find_max_mixture = 0
      this.round = 0
      for(var ax = 0; ax<this.sum_all_4_ttl_marker.length; ax++){
          if(this.round == 0){
          if(this.sum_all_4_ttl_marker[ax].value > 0){
            this.find_max_ttl_marker = this.sum_all_4_ttl_marker[ax].value
            this.find_max_ttl_marker_name = this.sum_all_4_ttl_marker[ax].value_name
            this.find_max_mixture = this.sum_all_4_ttl_marker[ax].value_mix
          }
          this.round ++
      }else{
        if(this.find_max_ttl_marker < this.sum_all_4_ttl_marker[ax].value){
            this.find_max_ttl_marker = this.sum_all_4_ttl_marker[ax].value
            this.find_max_ttl_marker_name = this.sum_all_4_ttl_marker[ax].value_name
            this.find_max_mixture = this.sum_all_4_ttl_marker[ax].value_mix
             
           }else if(this.find_max_ttl_marker == 0){
             this.find_max_ttl_marker = this.sum_all_4_ttl_marker[ax].value
            this.find_max_ttl_marker_name = this.sum_all_4_ttl_marker[ax].value_name
            this.find_max_mixture = this.sum_all_4_ttl_marker[ax].value_mix
           }
            
      }
      this.round ++
      }
      
      this.result_woven_diff1 = Number(this.end_loss_woven) - Number(this.actual_knit_woven)
      this.result_max_woven.push({
        value:this.result_woven_diff1,
        value_name1:"Woven",
        value_name2:"End Loss",
        value_target:this.end_loss_woven,
        value_actual:this.actual_knit_woven,
        number:1
      })
      
      

    
      this.result_woven_diff2 =  Number( this.splice_loss_woven) -Number(this.actual_knit_woven_splice)
     
      this.result_max_woven.push({
        value:this.result_woven_diff2,
        value_name1:"Woven",
        value_name2:"Splice Loss",
        value_target:this.splice_loss_woven,
        value_actual:this.actual_knit_woven_splice,
        number:1
      })
      
      
     
      this.result_woven_diff3 =  Number(this.cut_loss_woven) - Number(this.actual_knit_woven_cut)
      
      this.result_max_woven.push({
        value:this.result_woven_diff3,
        value_name1:"Woven",
        value_name2:"Cut Out Loss",
        value_target:this.cut_loss_woven,
        value_actual:this.actual_knit_woven_cut,
        number:1
      })
      
      
     
      this.result_woven_diff4 =  Number(this.rem_loss_woven) - Number(this.actual_knit_woven_rem)
      
       this.result_max_woven.push({
        value:this.result_woven_diff4,
        value_name1:"Woven",
        value_name2:"Remnants Loss",
        value_target:this.rem_loss_woven,
        value_actual:this.actual_knit_woven_rem,
        number:1
      })
      
      

      
     console.log(this.end_loss_light)
     console.log(this.actual_knit_light)
      this.result_light_diff1 = Number(this.end_loss_light) - Number(this.actual_knit_light)
     
      this.result_max_light.push({
        value:this.result_light_diff1,
        value_name1:"Light & Mid.Wt",
        value_name2:"End Loss",
        value_target:this.end_loss_light,
        value_actual:this.actual_knit_light,
        number:2
      })
      
      

      this.result_light_diff2 = Number(this.splice_loss_light) - Number(this.actual_knit_light_splice)
      
      this.result_max_light.push({
        value:this.result_light_diff2,
        value_name1:"Light & Mid.Wt",
        value_name2:"Splice Loss",
        value_target:this.splice_loss_light,
        value_actual:this.actual_knit_light_splice,
        number:2
      })
      
      
     
      this.result_light_diff3 = Number(this.cut_loss_light) - Number(this.actual_knit_light_cut)
     
      this.result_max_light.push({
        value:this.result_light_diff3,
        value_name1:"Light & Mid.Wt",
        value_name2:"Cut Out Loss",
        value_target:this.cut_loss_light,
        value_actual:this.actual_knit_light_cut,
        number:2
      })
      
      

   
      this.result_light_diff4 = Number(this.rem_loss_light) - Number(this.actual_knit_light_rem)
       
      this.result_max_light.push({
        value:this.result_light_diff4,
        value_name1:"Light & Mid.Wt",
        value_name2:"Cut Out Loss",
        value_target:this.rem_loss_light,
        value_actual:this.actual_knit_light_rem,
        number:2
      })
       
      
      
      this.result_fleece_diff1 = Number(this.end_loss_fleece) - Number(this.actual_knit_fleece)
      
      this.result_max_fleece.push({
        value:this.result_fleece_diff1,
        value_name1:"Fleece",
        value_name2:"End Loss",
        value_target:this.end_loss_fleece,
        value_actual:this.actual_knit_fleece,
        number:3
      })
      
      
     
      this.result_fleece_diff2 = Number(this.splice_loss_fleece) - Number(this.actual_knit_fleece_splice)
     
      this.result_max_fleece.push({
        value:this.result_fleece_diff2,
        value_name1:"Fleece",
        value_name2:"Splice Loss",
        value_target:this.splice_loss_fleece,
        value_actual:this.actual_knit_fleece_splice,
        number:3
      })
      
      
      
      this.result_fleece_diff3 = Number(this.cut_loss_fleece) - Number(this.actual_knit_fleece_cut)
     
      this.result_max_fleece.push({
        value:this.result_fleece_diff3,
        value_name1:"Fleece",
        value_name2:"Cut Out Loss",
        value_target:this.cut_loss_fleece,
        value_actual:this.actual_knit_fleece_cut,
        number:3
      })
      
      
      
      this.result_fleece_diff4 = Number(this.rem_loss_fleece) - Number(this.actual_knit_fleece_rem)
      
      this.result_max_fleece.push({
        value:this.result_fleece_diff4,
        value_name1:"Fleece",
        value_name2:"Remnants Loss",
        value_target:this.rem_loss_fleece,
        value_actual:this.actual_knit_fleece_rem,
        number:3
      })
      
      
      
      this.result_special_diff1 = Number(this.end_loss_special) - Number(this.actual_knit_special)
      
       this.result_max_special.push({
        value:this.result_special_diff1,
        value_name1:"Special",
        value_name2:"End Loss",
        value_target:this.end_loss_special,
        value_actual:this.actual_knit_special,
        number:4
      })
      
      
      
      this.result_special_diff2 = Number(this.splice_loss_special) - Number(this.actual_knit_special_splice)
      
      this.result_max_special.push({
        value:this.result_special_diff2,
        value_name1:"Special",
        value_name2:"Splice Loss",
        value_target:this.splice_loss_special,
        value_actual:this.actual_knit_special_splice,
        number:4
      })
      
      
      
      this.result_special_diff3 = Number(this.cut_loss_special) - Number(this.actual_knit_special_cut)
      
      this.result_max_special.push({
        value:this.result_special_diff3,
        value_name1:"Special",
        value_name2:"Cut Out Loss",
        value_target:this.cut_loss_special,
        value_actual:this.actual_knit_special_cut,
        number:4
      })
      
      
      
      this.result_special_diff4 = Number(this.rem_loss_special) - Number(this.actual_knit_special_rem)

      this.result_max_special.push({
        value:this.result_special_diff4,
        value_name1:"Special",
        value_name2:"Remnants Loss",
        value_target:this.rem_loss_special,
        value_actual:this.actual_knit_special_rem,
        number:4
      })

      
      this.status_check_number_woven = true;
      this.count_woven = 0
      for(var ax =  0; ax<this.result_max_woven.length; ax++){
          if(this.result_max_woven[ax].value >  0 || this.result_max_woven[ax].value == 0){
            
          }else if(this.result_max_woven[ax].value <  0){
               this.count_woven = Number(this.count_woven)+1
          }
      }
      if(this.count_woven > 0){
        this.status_check_number_woven = false;
      }
      
      this.find_max_4_woven_value = 0
      this.find_max_4_woven_name1 = ""
      this.find_max_4_woven_name2 = ""
      this.find_max_4_woven_target = 0
      this.find_max_4_woven_actual = 0
      this.find_max_woven_number = 0
      this.round = 0

      if(this.status_check_number_woven == true){
      for(var ax = 0; ax<this.result_max_woven.length; ax++){
          
          if(this.round == 0){
        
          if(this.result_max_woven[ax].value_actual != 0 ){
            this.find_max_4_woven_value = this.result_max_woven[ax].value
            this.find_max_4_woven_name1 = this.result_max_woven[ax].value_name1
            this.find_max_4_woven_name2 = this.result_max_woven[ax].value_name2
            this.find_max_4_woven_target = this.result_max_woven[ax].value_target
            this.find_max_4_woven_actual = this.result_max_woven[ax].value_actual
            this.find_max_woven_number = this.result_max_woven[ax].number
          }
          this.round ++
      }else{
         if(this.result_max_woven[ax].value_actual != 0 ){
        if(this.find_max_4_woven_value < this.result_max_woven[ax].value){
            this.find_max_4_woven_value = this.result_max_woven[ax].value
            this.find_max_4_woven_name1 = this.result_max_woven[ax].value_name1
            this.find_max_4_woven_name2 = this.result_max_woven[ax].value_name2
            this.find_max_4_woven_target = this.result_max_woven[ax].value_target
            this.find_max_4_woven_actual = this.result_max_woven[ax].value_actual
            this.find_max_woven_number = this.result_max_woven[ax].number
           }
         }
            
      }
      this.round ++
      }
      }else{
        for(var ax = 0; ax<this.result_max_woven.length; ax++){
          if(this.round2 == 0){
          if(this.result_max_woven[ax].value_actual  != 0){
            this.find_max_4_woven_value = this.result_max_woven[ax].value
            this.find_max_4_woven_name1 = this.result_max_woven[ax].value_name1
            this.find_max_4_woven_name2 = this.result_max_woven[ax].value_name2
            this.find_max_4_woven_target = this.result_max_woven[ax].value_target
            this.find_max_4_woven_actual = this.result_max_woven[ax].value_actual
            this.find_max_woven_number = this.result_max_woven[ax].number
          }
          this.round ++
      }else{
        if(this.result_max_woven[ax].value_actual != 0 ){
          if(this.result_max_woven[ax].value < 0){
           if(this.find_max_4_woven_value > this.result_max_woven[ax].value){
            this.find_max_4_woven_value = this.result_max_woven[ax].value
            this.find_max_4_woven_name1 = this.result_max_woven[ax].value_name1
            this.find_max_4_woven_name2 = this.result_max_woven[ax].value_name2
            this.find_max_4_woven_target = this.result_max_woven[ax].value_target
            this.find_max_4_woven_actual = this.result_max_woven[ax].value_actual
            this.find_max_woven_number = this.result_max_woven[ax].number
           }

          }
        }   
      }
      this.round ++
      }
      }
       this.minus_value_actual_woven = 0
       this.minus_value_actual_woven = Number(this.total_actual_knit_woven) - Number(this.Woven_total)
      this.result_per_week.push({
         value:this.find_max_4_woven_value,
        value_name1:this.find_max_4_woven_name1,
        value_name2:this.find_max_4_woven_name2,
        value_target:this.find_max_4_woven_target,
        value_actual:this.find_max_4_woven_actual,
        number:this.find_max_woven_number,
        all_value_actual:this.minus_value_actual_woven
      })


      this.status_check_number_light = true;
      this.count_light = 0
      for(var ax =  0; ax<this.result_max_light.length; ax++){
          if(this.result_max_light[ax].value >  0 || this.result_max_light[ax].value == 0){
            
          }else if(this.result_max_light[ax].value <  0){
               this.count_light = Number(this.count_light)+1
          }
      }
      
      if(this.count_light > 0){
        this.status_check_number_light = false;
      }
   
      this.find_max_4_light_value = 0
      this.find_max_4_light_name1 = ""
      this.find_max_4_light_name2 = ""
      this.find_max_4_light_target = 0
      this.find_max_4_light_actual = 0
      this.find_max_light_number = 0
      this.round2 = 0
      if(this.status_check_number_light == true){

      }else{
      for(var ax = 0; ax<this.result_max_light.length; ax++){
          if(this.round2 == 0){
          if(this.result_max_light[ax].value_actual  != 0){
            this.find_max_4_light_value = this.result_max_light[ax].value
            this.find_max_4_light_name1 = this.result_max_light[ax].value_name1
            this.find_max_4_light_name2 = this.result_max_light[ax].value_name2
            this.find_max_4_light_target = this.result_max_light[ax].value_target
            this.find_max_4_light_actual = this.result_max_light[ax].value_actual
            this.find_max_light_number = this.result_max_light[ax].number
          }
          this.round2 ++
      }else{
        if(this.result_max_light[ax].value_actual != 0 ){
          if(this.result_max_light[ax].value < 0){
           if(this.find_max_4_light_value > this.result_max_light[ax].value){
            this.find_max_4_light_value = this.result_max_light[ax].value
            this.find_max_4_light_name1 = this.result_max_light[ax].value_name1
            this.find_max_4_light_name2 = this.result_max_light[ax].value_name2
            this.find_max_4_light_target = this.result_max_light[ax].value_target
            this.find_max_4_light_actual = this.result_max_light[ax].value_actual
            this.find_max_light_number = this.result_max_light[ax].number
           }

          }
        }   
      }
      this.round2 ++
      }
      }
      this.minus_value_actual_light = 0
       this.minus_value_actual_light = Number(this.total_actual_knit_light) - Number(this.Light_total)
      this.result_per_week.push({
        value:this.find_max_4_light_value,
        value_name1:this.find_max_4_light_name1,
        value_name2:this.find_max_4_light_name2,
        value_target:this.find_max_4_light_target,
        value_actual:this.find_max_4_light_actual,
        number:this.find_max_light_number,
        all_value_actual:this.minus_value_actual_light
      })

      this.status_check_number_fleece = true;
      this.count_fleece = 0
      for(var ax =  0; ax<this.result_max_fleece.length; ax++){
          if(this.result_max_fleece[ax].value >  0 || this.result_max_fleece[ax].value == 0){
            
          }else if(this.result_max_fleece[ax].value <  0){
               this.count_fleece = Number(this.count_fleece)+1
          }
      }
      if(this.count_fleece > 0){
        this.status_check_number_fleece = false;
      }
      this.find_max_4_fleece_value = 0
      this.find_max_4_fleece_name1 = ""
      this.find_max_4_fleece_name2 = ""
      this.find_max_4_fleece_target = 0
      this.find_max_4_fleece_actual = 0
      this.find_max_fleece_number = 0
      this.round3 = 0
      if(this.status_check_number_fleece == true){
      for(var ax = 0; ax<this.result_max_fleece.length; ax++){
          if(this.round3 == 0){
          if(this.result_max_fleece[ax].value_actual  != 0){
            this.find_max_4_fleece_value = this.result_max_fleece[ax].value
            this.find_max_4_fleece_name1 = this.result_max_fleece[ax].value_name1
            this.find_max_4_fleece_name2 = this.result_max_fleece[ax].value_name2
            this.find_max_4_fleece_target = this.result_max_fleece[ax].value_target
            this.find_max_4_fleece_actual = this.result_max_fleece[ax].value_actual
            this.find_max_fleece_number = this.result_max_fleece[ax].number
          }
          this.round3 ++
      }else{
        if(this.result_max_fleece[ax].value_actual  != 0){
        if(this.find_max_4_fleece_value < this.result_max_fleece[ax].value){
            this.find_max_4_fleece_value = this.result_max_fleece[ax].value
            this.find_max_4_fleece_name1 = this.result_max_fleece[ax].value_name1
            this.find_max_4_fleece_name2 = this.result_max_fleece[ax].value_name2
            this.find_max_4_fleece_target = this.result_max_fleece[ax].value_target
            this.find_max_4_fleece_actual = this.result_max_fleece[ax].value_actual
            this.find_max_fleece_number = this.result_max_fleece[ax].number
           }
        }  
      }
      this.round3 ++
      }
      }else{
          for(var ax = 0; ax<this.result_max_fleece.length; ax++){
          if(this.round2 == 0){
          if(this.result_max_fleece[ax].value_actual  != 0){
            this.find_max_4_fleece_value = this.result_max_fleece[ax].value
            this.find_max_4_fleece_name1 = this.result_max_fleece[ax].value_name1
            this.find_max_4_fleece_name2 = this.result_max_fleece[ax].value_name2
            this.find_max_4_fleece_target = this.result_max_fleece[ax].value_target
            this.find_max_4_fleece_actual = this.result_max_fleece[ax].value_actual
            this.find_max_fleece_number = this.result_max_fleece[ax].number
          }
          this.round3 ++
      }else{
        if(this.result_max_fleece[ax].value_actual != 0 ){
          if(this.result_max_fleece[ax].value < 0){
           if(this.find_max_4_fleece_value > this.result_max_fleece[ax].value){
            this.find_max_4_fleece_value = this.result_max_fleece[ax].value
            this.find_max_4_fleece_name1 = this.result_max_fleece[ax].value_name1
            this.find_max_4_fleece_name2 = this.result_max_fleece[ax].value_name2
            this.find_max_4_fleece_target = this.result_max_fleece[ax].value_target
            this.find_max_4_fleece_actual = this.result_max_fleece[ax].value_actual
            this.find_max_fleece_number = this.result_max_fleece[ax].number
           }

          }
        }   
      }
      this.round3 ++
      }
      }
      this.minus_value_actual_fleece = 0
       this.minus_value_actual_fleece = Number(this.total_actual_knit_fleece) - Number(this.Fleece_total)
      this.result_per_week.push({
         value:this.find_max_4_fleece_value,
        value_name1:this.find_max_4_fleece_name1,
        value_name2:this.find_max_4_fleece_name2,
        value_target:this.find_max_4_fleece_target,
        value_actual:this.find_max_4_fleece_actual,
        number:this.find_max_fleece_number,
        all_value_actual:this.minus_value_actual_fleece
      })


      this.status_check_number_special = true;
      this.count_special = 0
      for(var ax =  0; ax<this.result_max_special.length; ax++){
          if(this.result_max_special[ax].value >  0 || this.result_max_special[ax].value == 0){
            
          }else if(this.result_max_special[ax].value <  0){
               this.count_special = Number(this.count_special)+1
          }
      }
      if(this.count_special > 0){
        this.status_check_number_special = false;
      }

      this.find_max_4_special_value = 0
      this.find_max_4_special_name1 = ""
      this.find_max_4_special_name2 = ""
      this.find_max_4_special_target = 0
      this.find_max_4_special_actual = 0
      this.find_max_special_number = 0
      this.round4 = 0
      if(this.status_check_number_special == true){
      for(var ax = 0; ax<this.result_max_special.length; ax++){
          if(this.round4 == 0){
          if(this.result_max_special[ax].value_actual  != 0){
            this.find_max_4_special_value = this.result_max_special[ax].value
            this.find_max_4_special_name1 = this.result_max_special[ax].value_name1
            this.find_max_4_special_name2 = this.result_max_special[ax].value_name2
            this.find_max_4_special_target = this.result_max_special[ax].value_target
            this.find_max_4_special_actual = this.result_max_special[ax].value_actual
            this.find_max_special_number = this.result_max_special[ax].number
          }
          this.round4 ++
      }else{
        if(this.result_max_special[ax].value_actual  != 0){
        if(this.find_max_4_special_value < this.result_max_special[ax].value){
            this.find_max_4_special_value = this.result_max_special[ax].value
            this.find_max_4_special_name1 = this.result_max_special[ax].value_name1
            this.find_max_4_special_name2 = this.result_max_special[ax].value_name2
            this.find_max_4_special_target = this.result_max_special[ax].value_target
            this.find_max_4_special_actual = this.result_max_special[ax].value_actual
            this.find_max_special_number = this.result_max_special[ax].number
           }
        }
            
      }
      this.round4 ++
      }
      }else{
        for(var ax = 0; ax<this.result_max_special.length; ax++){
          if(this.round2 == 0){
          if(this.result_max_special[ax].value_actual  != 0){
            this.find_max_4_special_value = this.result_max_special[ax].value
            this.find_max_4_special_name1 = this.result_max_special[ax].value_name1
            this.find_max_4_special_name2 = this.result_max_special[ax].value_name2
            this.find_max_4_special_target = this.result_max_special[ax].value_target
            this.find_max_4_special_actual = this.result_max_special[ax].value_actual
            this.find_max_special_number = this.result_max_special[ax].number
          }
          this.round3 ++
      }else{
        if(this.result_max_special[ax].value_actual != 0 ){
          if(this.result_max_special[ax].value < 0){
           if(this.find_max_4_special_value > this.result_max_special[ax].value){
            this.find_max_4_special_value = this.result_max_special[ax].value
            this.find_max_4_special_name1 = this.result_max_special[ax].value_name1
            this.find_max_4_special_name2 = this.result_max_special[ax].value_name2
            this.find_max_4_special_target = this.result_max_special[ax].value_target
            this.find_max_4_special_actual = this.result_max_special[ax].value_actual
            this.find_max_special_number = this.result_max_special[ax].number
           }

          }
        }   
      }
      this.round3 ++
      }
      }
     this.minus_value_actual_special = 0
       this.minus_value_actual_special = Number(this.total_actual_knit_special) - Number(this.Special_total)
      this.result_per_week.push({
        value:this.find_max_4_special_value,
        value_name1:this.find_max_4_special_name1,
        value_name2:this.find_max_4_special_name2,
        value_target:this.find_max_4_special_target,
        value_actual:this.find_max_4_special_actual,
        number:this.find_max_special_number,
        all_value_actual:this.minus_value_actual_special
      })

    console.log(this.result_per_week)
    

      this.status_check_number_per_week = true;
      this.count_per_week = 0
      for(var ax =  0; ax<this.result_per_week.length; ax++){
          if(this.result_per_week[ax].value >  0 || this.result_per_week[ax].value == 0){
            
          }else if(this.result_per_week[ax].value <  0){
               this.count_per_week = Number(this.count_per_week)+1
          }
      }
      if(this.count_per_week > 0){
        this.status_check_number_per_week = false;
      }

      this.find_max_4_all_value = 0
      this.find_max_4_all_name1 = ""
      this.find_max_4_all_name2 = ""
      this.find_max_4_all_target = 0
      this.find_max_4_all_actual = 0
      this.find_max_all_number = 0
      this.find_max_minus_all_actual = 0 //ผล - ของค่า actual กับ Target
      this.round5 = 0
   
      for(var ax = 0; ax<this.result_per_week.length; ax++){
          if(this.round5 == 0){
        
            this.find_max_4_all_value = this.result_per_week[ax].value
            this.find_max_4_all_name1 = this.result_per_week[ax].value_name1
            this.find_max_4_all_name2 = this.result_per_week[ax].value_name2
            this.find_max_4_all_target = this.result_per_week[ax].value_target
            this.find_max_4_all_actual = this.result_per_week[ax].value_actual
            this.find_max_all_number = this.result_per_week[ax].number
            this.find_max_minus_all_actual = this.result_per_week[ax].all_value_actual
          
           
          
           this.round5 ++
          
      }else if(this.round5 > 0){
       
            if(this.find_max_minus_all_actual < this.result_per_week[ax].all_value_actual){
            this.find_max_4_all_value = this.result_per_week[ax].value
            this.find_max_4_all_name1 = this.result_per_week[ax].value_name1
            this.find_max_4_all_name2 = this.result_per_week[ax].value_name2
            this.find_max_4_all_target = this.result_per_week[ax].value_target
            this.find_max_4_all_actual = this.result_per_week[ax].value_actual
            this.find_max_all_number = this.result_per_week[ax].number
            this.find_max_minus_all_actual = this.result_per_week[ax].all_value_actual
            }
           
          
      }
    
      }
    
    
      
   

    if(this.find_max_4_all_value < 0){
      this.find_max_4_all_value = 0-(this.find_max_4_all_value)
    }
    

      


      worksheet.getCell("R15").value = "สรุปผลการดำเนินการสัปดาห์ปัจจุบัน\nFabric Type"+this.find_max_all_number+"("+this.find_max_4_all_name1+")"+"สูงสุดมาจาก\n"+this.find_max_4_all_name2 +"\nTarget = "+parseFloat(this.find_max_4_all_target).toFixed(2) +"%\nActual = "+parseFloat(this.find_max_4_all_actual).toFixed(2)+"%\nDiff. = "+parseFloat(this.find_max_4_all_value).toFixed(2)+"%\nปริมาณการใช้ผ้าสูงสุด\nFabric Diff("+this.find_max_ttl_marker_name+")"+"  "+parseFloat(this.find_max_mixture).toFixed(2)+"%";

       worksheet.mergeCells("R15:S31");
       
          worksheet.getCell("R15").font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
        worksheet.getCell("R15").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
       worksheet.getCell("R15").alignment = {
          vertical: "top",
          horizontal: "center",
          wrapText: true
        };
     
      


         worksheet.getCell("BX43").value = "SumType 1";
     
       worksheet.mergeCells("BX43:BX44");


       /*  this.ttl_marker_endless_woven.push({
              end_ttl_marker_woven : this.ttl_marker_end_woven
            }) */
            
       //summary type 1
       
          this.summary_ttl_marker_end_woven = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary_ttl_marker_end_woven = Number(this.summary_ttl_marker_end_woven) + Number(this.ttl_marker_endless_woven[ax].end_ttl_marker_woven)
            }
        
          

         worksheet.getCell("BX45").value = Number(this.summary_ttl_marker_end_woven)
         worksheet.getCell("BX45").numFmt = '0.00'; 

      
   

            this.summary_ttl_marker_endloss_woven = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary_ttl_marker_endloss_woven = Number(this.summary_ttl_marker_endloss_woven) + Number(this.ttl_endloss_yd_woven[ax].end_loss_yd_woven)
            } 

       worksheet.getCell("BX46").value = Number(this.summary_ttl_marker_endloss_woven)
       worksheet.getCell("BX46").numFmt = '0.00';


        /*  this.ttl_splice_woven.push({
              splice_loss_woven : this.splice_woven
            }) */

         this.summary_ttl_marker_splice_woven = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary_ttl_marker_splice_woven = Number(this.summary_ttl_marker_splice_woven) + Number(this.ttl_splice_woven[ax].splice_loss_woven)
            } 


          worksheet.getCell("BX47").value = Number(this.summary_ttl_marker_splice_woven);
          worksheet.getCell("BX47").numFmt = '0.00';
  


      worksheet.getCell("BX63").value = "SumType 2";
       worksheet.mergeCells("BX63:BX64");
  

     
         this.summary_ttl_marker_cut_woven = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary_ttl_marker_cut_woven = Number(this.summary_ttl_marker_cut_woven) + Number(this.ttl_cut_woven[ax].cut_loss_woven)
            } 

          worksheet.getCell("BX48").value = Number(this.summary_ttl_marker_cut_woven)
          worksheet.getCell("BX48").numFmt = '0.00';
  

        this.summary_ttl_marker_rem_woven = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary_ttl_marker_rem_woven = Number(this.summary_ttl_marker_rem_woven) + Number(this.ttl_rem_woven[ax].rem_loss_woven)
            } 
         worksheet.getCell("BX49").value = Number(this.summary_ttl_marker_rem_woven)
         worksheet.getCell("BX49").numFmt = '0.00';
  

      /*   this.ttl_woven.push({
             value_all_woven : this.all_ttl_act
           }) */
        this.summary_ttl_marker_all_woven = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary_ttl_marker_all_woven = Number(this.summary_ttl_marker_all_woven) + Number(this.ttl_woven[ax].value_all_woven)
            } 
         worksheet.getCell("BX50").value = Number(this.summary_ttl_marker_all_woven)
         worksheet.getCell("BX50").numFmt = '0.00';
              
     //summary type 2
      this.summary2_ttl_marker_end_light = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary2_ttl_marker_end_light = Number(this.summary2_ttl_marker_end_light) + Number(this.ttl_marker_endless_light[ax].end_ttl_marker_light)
            } 
        worksheet.getCell("BX65").value = Number(this.summary2_ttl_marker_end_light)
        worksheet.getCell("BX65").numFmt = '0.00';
     

        /* this.ttl_endloss_yd_light.push({
              end_loss_yd_light : this.end_loss_yd_light
            }) */
      

        this.summary2_ttl_marker_endless_light = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary2_ttl_marker_endless_light = Number(this.summary2_ttl_marker_endless_light) + Number(this.ttl_endloss_yd_light[ax].end_loss_yd_light)
            } 
        worksheet.getCell("BX66").value = Number(this.summary2_ttl_marker_endless_light)
        worksheet.getCell("BX66").numFmt = '0.00';
 

     /*    this.ttl_splice_light.push({
              splice_loss_light : 0
        })
 */
        this.summary2_ttl_marker_splice_light = 0    
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary2_ttl_marker_splice_light = Number(this.summary2_ttl_marker_splice_light) + Number(this.ttl_splice_light[ax].splice_loss_light)
            } 

          worksheet.getCell("BX67").value = Number(this.summary2_ttl_marker_splice_light)
          worksheet.getCell("BX67").numFmt = '0.00';
  

        /*    this.ttl_cut_light.push({
              cut_loss_light : 0
        })
 */
        this.summary2_ttl_marker_cut_light = 0    
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary2_ttl_marker_cut_light = Number(this.summary2_ttl_marker_cut_light) + Number(this.ttl_cut_light[ax].cut_loss_light)
            } 


         worksheet.getCell("BX68").value = Number(this.summary2_ttl_marker_cut_light)
         worksheet.getCell("BX68").numFmt = '0.00';
 


        this.summary2_ttl_marker_rem_light = 0    
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary2_ttl_marker_rem_light = Number(this.summary2_ttl_marker_rem_light) + Number(this.ttl_rem_light[ax].rem_loss_light)
            } 

          worksheet.getCell("BX69").value = Number(this.summary2_ttl_marker_rem_light)
          worksheet.getCell("BX69").numFmt = '0.00';
    

        /*   this.ttl_light.push({
             value_all_light : this.all_ttl_act
           }) */
        this.summary_ttl_marker_all_light = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary_ttl_marker_all_light = Number(this.summary_ttl_marker_all_light) + Number(this.ttl_light[ax].value_all_light)
            } 
         worksheet.getCell("BX70").value = Number(this.summary_ttl_marker_all_light)
         worksheet.getCell("BX70").numFmt = '0.00';
     

         worksheet.getCell("BX84").value = "SumType 3";
       worksheet.mergeCells("BX84:BX85");
      


         //summary type 3
      this.summary3_ttl_marker_end_fleece = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary3_ttl_marker_end_fleece = Number(this.summary3_ttl_marker_end_fleece) + Number(this.ttl_marker_endless_fleece[ax].end_ttl_marker_fleece)
            } 
        worksheet.getCell("BX86").value = Number(this.summary3_ttl_marker_end_fleece)
        worksheet.getCell("BX86").numFmt = '0.00';
   

        this.summary3_ttl_marker_endless_fleece = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary3_ttl_marker_endless_fleece = Number(this.summary3_ttl_marker_endless_fleece) + Number(this.ttl_endloss_yd_fleece[ax].end_loss_yd_fleece)
            } 
        worksheet.getCell("BX87").value = Number(this.summary3_ttl_marker_endless_fleece)
        worksheet.getCell("BX87").numFmt = '0.00';
   

         this.summary3_ttl_marker_splice_fleece = 0    
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary3_ttl_marker_splice_fleece = Number(this.summary3_ttl_marker_splice_fleece) + Number(this.ttl_splice_fleece[ax].splice_loss_fleece)
            } 

          worksheet.getCell("BX88").value = Number(this.summary3_ttl_marker_splice_fleece)
          worksheet.getCell("BX88").numFmt = '0.00';
    

         this.summary3_ttl_marker_cut_fleece = 0    
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary3_ttl_marker_cut_fleece = Number(this.summary3_ttl_marker_cut_fleece) + Number(this.ttl_cut_fleece[ax].cut_loss_fleece)
            } 


         worksheet.getCell("BX89").value = Number(this.summary3_ttl_marker_cut_fleece)
         worksheet.getCell("BX89").numFmt = '0.00';
 

        this.summary3_ttl_marker_rem_fleece = 0    
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary3_ttl_marker_rem_fleece = Number(this.summary3_ttl_marker_rem_fleece) + Number(this.ttl_rem_fleece[ax].rem_loss_fleece)
            } 

          worksheet.getCell("BX90").value = Number(this.summary3_ttl_marker_rem_fleece);
          worksheet.getCell("BX90").numFmt = '0.00';
         

         this.summary3_ttl_marker_all_fleece = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary3_ttl_marker_all_fleece = Number(this.summary3_ttl_marker_all_fleece) + Number(this.ttl_fleece[ax].value_all_fleece)
            } 
         worksheet.getCell("BX91").value = Number(this.summary3_ttl_marker_all_fleece)
         worksheet.getCell("BX91").numFmt = '0.00';
   


        worksheet.getCell("BX104").value = "SumType 4";
       worksheet.mergeCells("BX104:BX105");
    


           //summary type 4
      this.summary4_ttl_marker_end_special = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary4_ttl_marker_end_special = Number(this.summary4_ttl_marker_end_special) + Number(this.ttl_marker_endless_special[ax].end_ttl_marker_special)
            } 
        worksheet.getCell("BX106").value = Number(this.summary4_ttl_marker_end_special)
        worksheet.getCell("BX106").numFmt = '0.00';
      

         this.summary4_ttl_marker_endless_special = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary4_ttl_marker_endless_special = Number(this.summary4_ttl_marker_endless_special) + Number(this.ttl_endloss_yd_special[ax].end_loss_yd_special)
            } 
        worksheet.getCell("BX107").value = Number(this.summary4_ttl_marker_endless_special)
        worksheet.getCell("BX107").numFmt = '0.00';
      


         this.summary4_ttl_marker_splice_special = 0    
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary4_ttl_marker_splice_special = Number(this.summary4_ttl_marker_splice_special) + Number(this.ttl_splice_special[ax].splice_loss_special)
            } 

          worksheet.getCell("BX108").value = Number(this.summary4_ttl_marker_splice_special)
          worksheet.getCell("BX108").numFmt = '0.00';
    


        this.summary4_ttl_marker_cut_special = 0    
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary4_ttl_marker_cut_special = Number(this.summary4_ttl_marker_cut_special) + Number(this.ttl_cut_special[ax].cut_loss_special)
            } 


         worksheet.getCell("BX109").value = Number(this.summary4_ttl_marker_cut_special)
         worksheet.getCell("BX109").numFmt = '0.00';
    

        this.summary4_ttl_marker_rem_special = 0    
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary4_ttl_marker_rem_special = Number(this.summary4_ttl_marker_rem_special) + Number(this.ttl_rem_special[ax].rem_loss_special)
            } 

          worksheet.getCell("BX110").value = Number(this.summary4_ttl_marker_rem_special);
          worksheet.getCell("BX110").numFmt = '0.00';
     

         this.summary4_ttl_marker_all_special = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary4_ttl_marker_all_special = Number(this.summary4_ttl_marker_all_special) + Number(this.ttl_special[ax].value_all_special)
            } 
         worksheet.getCell("BX111").value = Number(this.summary4_ttl_marker_all_special);
         worksheet.getCell("BX111").numFmt = '0.00';
      

        worksheet.getCell("BX123").value = "SumType 1-4";
     
       worksheet.mergeCells("BX123:BX124");
     

      /* this.total_ac_end_arr.push({
           ac_end : this.total_ac_end_loss_1_4
         }) */

       this.summary5_ttl_marker_all_end_1_4 = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary5_ttl_marker_all_end_1_4 = Number(this.summary5_ttl_marker_all_end_1_4) + Number(this.total_ac_end_arr[ax].ac_end)
            } 
         worksheet.getCell("BX125").value = Number(this.summary5_ttl_marker_all_end_1_4)
         worksheet.getCell("BX125").numFmt = '0.00';
    
        this.summary5_ttl_marker_all_splice_1_4 = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary5_ttl_marker_all_splice_1_4 = Number(this.summary5_ttl_marker_all_splice_1_4) + Number(this.total_ac_splice_arr[ax].ac_splice)
            } 
         worksheet.getCell("BX126").value = Number(this.summary5_ttl_marker_all_splice_1_4)
         worksheet.getCell("BX126").numFmt = '0.00';
      


          this.summary5_ttl_marker_all_cut_1_4 = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
            this.summary5_ttl_marker_all_cut_1_4 = Number(this.summary5_ttl_marker_all_cut_1_4) + Number(this.total_ac_cut_arr[ax].ac_cut)
            } 
         worksheet.getCell("BX127").value = Number(this.summary5_ttl_marker_all_cut_1_4)
         worksheet.getCell("BX127").numFmt = '0.00';
      


           this.summary5_ttl_marker_all_rem_1_4 = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
       
           this.summary5_ttl_marker_all_rem_1_4 = Number(this.summary5_ttl_marker_all_rem_1_4) + Number(this.total_ac_rem_arr[ax].ac_rem)
            } 
         worksheet.getCell("BX128").value = Number(this.summary5_ttl_marker_all_rem_1_4)
         worksheet.getCell("BX128").numFmt = '0.00';
 

        /*  this.total_ac_all_arr.push({
              value_all_ac : this.total_act_spread_yd_1_4
            })
 */

          this.summary5_ttl_marker_all_1_4 = 0
         for(var ax = 0; ax<this.ttl_marker_endless_woven.length; ax++){
       
           this.summary5_ttl_marker_all_1_4 = Number(this.summary5_ttl_marker_all_1_4) + Number(this.total_ac_all_arr[ax].value_all_ac)
            } 
         worksheet.getCell("BX129").value = Number(this.summary5_ttl_marker_all_1_4)
         worksheet.getCell("BX129").numFmt = '0.00';
   

         worksheet.getCell("X5").value = "Cumulate Type";


        worksheet.getCell("X6").value = 1;
     
      

       worksheet.getCell("X7").value = Number(this.summary_ttl_marker_endloss_woven)
       worksheet.getCell("X7").numFmt = '0.00';

        worksheet.getCell("X8").value = Number(this.summary_ttl_marker_splice_woven)
        worksheet.getCell("X8").numFmt = '0.00';


         worksheet.getCell("X9").value = Number(this.summary_ttl_marker_cut_woven)
         worksheet.getCell("X9").numFmt = '0.00';
     

          worksheet.getCell("X10").value = Number(this.summary_ttl_marker_rem_woven)
          worksheet.getCell("X10").numFmt = '0.00';
    

        this.cumulative_woven = 0
        this.cumulative_woven = Number(this.summary_ttl_marker_endloss_woven) + Number(this.summary_ttl_marker_splice_woven)
        +Number(this.summary_ttl_marker_cut_woven) + Number(this.summary_ttl_marker_rem_woven)
          worksheet.getCell("X11").value = Number(this.cumulative_woven)
          worksheet.getCell("X11").numFmt = '0.00';
     


        worksheet.getCell("Y6").value = 2;
    
         worksheet.getCell("Y7").value = Number(this.summary2_ttl_marker_endless_light)
         worksheet.getCell("Y7").numFmt = '0.00';
      

         worksheet.getCell("Y8").value = Number(this.summary2_ttl_marker_splice_light)
         worksheet.getCell("Y8").numFmt = '0.00';
   

         worksheet.getCell("Y9").value = Number(this.summary2_ttl_marker_cut_light)
         worksheet.getCell("Y9").numFmt = '0.00';
       

          worksheet.getCell("Y10").value = Number(this.summary2_ttl_marker_rem_light);
          worksheet.getCell("Y10").numFmt = '0.00';
      

        this.cumulative_light = 0

        this.cumulative_light = Number(this.summary2_ttl_marker_endless_light) + Number(this.summary2_ttl_marker_splice_light)
        +Number(this.summary2_ttl_marker_cut_light) + Number(this.summary2_ttl_marker_rem_light)
          worksheet.getCell("Y11").value = Number(this.cumulative_light)
          worksheet.getCell("Y11").numFmt = '0.00';
      

       


        worksheet.getCell("Z6").value = 3;
     

        worksheet.getCell("Z7").value = Number(this.summary3_ttl_marker_endless_fleece)
        worksheet.getCell("Z7").numFmt = '0.00';

    

        worksheet.getCell("Z8").value = Number(this.summary3_ttl_marker_splice_fleece)
        worksheet.getCell("Z8").numFmt = '0.00';
     

       

        worksheet.getCell("Z9").value = Number(this.summary3_ttl_marker_cut_fleece)
        worksheet.getCell("Z9").numFmt = '0.00';
      


        worksheet.getCell("Z10").value = Number(this.summary3_ttl_marker_rem_fleece)
        worksheet.getCell("Z10").numFmt = '0.00';
      


        
           this.cumulative_fleece = 0
        this.cumulative_fleece = Number(this.summary3_ttl_marker_endless_fleece) + Number(this.summary3_ttl_marker_splice_fleece)
        +Number(this.summary3_ttl_marker_cut_fleece) + Number(this.summary3_ttl_marker_rem_fleece)
          worksheet.getCell("Z11").value = Number(this.cumulative_fleece)
          worksheet.getCell("Z11").numFmt = '0.00';

          worksheet.getCell("AA6").value = 4;
   

          worksheet.getCell("AA7").value = Number(this.summary4_ttl_marker_endless_special)
          worksheet.getCell("AA7").numFmt = '0.00';
   

          worksheet.getCell("AA8").value = Number(this.summary4_ttl_marker_splice_special)
          worksheet.getCell("AA8").numFmt = '0.00';
  


          worksheet.getCell("AA9").value = Number(this.summary4_ttl_marker_cut_special)
          worksheet.getCell("AA9").numFmt = '0.00';


          worksheet.getCell("AA10").value = Number(this.summary4_ttl_marker_rem_special)
          worksheet.getCell("AA10").numFmt = '0.00';

               this.cumulative_special = 0
        this.cumulative_special = Number(this.summary4_ttl_marker_endless_special) + Number(this.summary4_ttl_marker_splice_special)
        +Number(this.summary4_ttl_marker_cut_special) + Number(this.summary4_ttl_marker_rem_special)
          worksheet.getCell("AA11").value = Number(this.cumulative_special)
          worksheet.getCell("AA11").numFmt = '0.00';
 

          worksheet.getCell("AB6").value = "SUM";
   

        this.cumulative_end = 0
        this.cumulative_end = Number(this.summary_ttl_marker_endloss_woven)+Number(this.summary2_ttl_marker_endless_light)
        +Number(this.summary3_ttl_marker_endless_fleece)+Number(this.summary4_ttl_marker_endless_special)
          worksheet.getCell("AB7").value = Number(this.cumulative_end)
          worksheet.getCell("AB7").numFmt = '0.00';
    


        this.cumulative_splice = 0
        this.cumulative_splice = Number(this.summary_ttl_marker_splice_woven)+Number(this.summary2_ttl_marker_splice_light)
        +Number(this.summary3_ttl_marker_splice_fleece)+Number(this.summary4_ttl_marker_splice_special)
          worksheet.getCell("AB8").value = Number(this.cumulative_splice)
          worksheet.getCell("AB8").numFmt = '0.00';
     


        this.cumulative_cut = 0
        this.cumulative_cut = Number(this.summary_ttl_marker_cut_woven)+Number(this.summary2_ttl_marker_cut_light)
        +Number(this.summary3_ttl_marker_cut_fleece)+Number(this.summary4_ttl_marker_cut_special)
          worksheet.getCell("AB9").value = Number(this.cumulative_cut)
          worksheet.getCell("AB9").numFmt = '0.00';
  

        this.cumulative_rem = 0
        this.cumulative_rem = Number(this.summary_ttl_marker_rem_woven)+Number(this.summary2_ttl_marker_rem_light)
        +Number(this.summary3_ttl_marker_rem_fleece)+Number(this.summary4_ttl_marker_rem_special)
          worksheet.getCell("AB10").value = Number(this.cumulative_rem)
          worksheet.getCell("AB10").numFmt = '0.00';
   


        this.cumulative_all = 0
        this.cumulative_all = Number(this.cumulative_woven)+Number(this.cumulative_light)
        +Number(this.cumulative_fleece)+Number(this.cumulative_special)
          worksheet.getCell("AB11").value = Number(this.cumulative_all)
         worksheet.getCell("AB11").numFmt = '0.00';
   

          worksheet.getCell("S13").value = Number(this.cumulative_all)
           worksheet.getCell("S13").numFmt = '#,##0.00';

      console.log(this.cumulative_end)
      console.log(this.all_sum_cal)
      this.cumulated_avg_end = 0
      this.cumulated_avg_end = this.cumulative_end/this.all_sum_cal*100
      

         worksheet.getCell("R7").value = this.cumulated_avg_end/100
         worksheet.getCell("R7").numFmt = '0.00%';
   

        this.cumulated_avg_splice = 0
      this.cumulated_avg_splice = this.cumulative_splice/this.all_sum_cal*100
      

         worksheet.getCell("R8").value = this.cumulated_avg_splice/100
          worksheet.getCell("R8").numFmt = '0.00%';
    


         this.cumulated_avg_cut = 0
      this.cumulated_avg_cut = this.cumulative_cut/this.all_sum_cal*100
      

         worksheet.getCell("R9").value = this.cumulated_avg_cut/100;
          worksheet.getCell("R9").numFmt = '0.00%';
 

          this.cumulated_avg_rem = 0
      this.cumulated_avg_rem = this.cumulative_rem/this.all_sum_cal*100
      

         worksheet.getCell("R10").value = this.cumulated_avg_rem/100;
          worksheet.getCell("R10").numFmt = '0.00%';
      
     

       this.cumulated_avg_all = 0
      this.cumulated_avg_all = this.cumulative_all/this.all_sum_cal*100
      
 
         worksheet.getCell("R11").value = this.cumulated_avg_all/100
         worksheet.getCell("R11").numFmt = '0.00%';



         worksheet.getCell("W152").value = "KPI แผนกตัด"
         worksheet.getCell("W153").value = "End Loss VS Target%"
         worksheet.getCell("W154").value = "Width Loss VS Target%"
         worksheet.getCell("W155").value = "Work Week Production"
         worksheet.getCell("W156").value = "End & Width Loss Avg. %"
         worksheet.getCell("W157").value = "Monthly"

       /*  this.last_total_spread_end.push({
            value_last_spread : this.total_end_target
          }) 

          this.arr_end_loss_1_4.push({
            value:this.end_loss_special_per
          }) 
 */
         
         
         
        for(var ax = 0; ax<this.arr_colume.length; ax++){
         this.end_loss_target_per = 0
          if(isNaN(this.arr_end_loss_1_4[ax].value)==true){
            this.arr_end_loss_1_4[ax].value = 0
          }
          if(isNaN(this.last_total_spread_end[ax].value_last_spread)==true){
            this.last_total_spread_end[ax].value_last_spread = 0
          }
         
          this.end_loss_target_per = this.last_total_spread_end[ax].value_last_spread / this.arr_end_loss_1_4[ax].value*100
          this.end_loss_target_arr.push({
            value:this.end_loss_target_per
          })
         
          
           if(isNaN(this.end_loss_target_per)==true || isFinite(this.end_loss_target_per)==false){
          worksheet.getCell(this.arr_colume[ax].name_col+"153").value = 0.00/100
          worksheet.getCell(this.arr_colume[ax].name_col+"153").numFmt = '0.00%';  
           }else{
          worksheet.getCell(this.arr_colume[ax].name_col+"153").value = this.end_loss_target_per/100
          worksheet.getCell(this.arr_colume[ax].name_col+"153").numFmt = '0.00%';
           }
        }

        
      

      

        for(var ax = 0; ax<this.sum_ttl_marker.length; ax++){
        
          this.ac_width = 0
          
          this.ac_width = this.sum_widthloss[ax].value / this.sum_ttl_marker[ax].value * 100
          
          
          this.ac_width_kpi = 1.50 / this.ac_width *100

          if(isNaN(this.ac_width_kpi)==true || isFinite(this.ac_width_kpi)==false || this.ac_width <= 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"154").value = 0.00/100
          worksheet.getCell(this.arr_colume[ax].name_col+"154").numFmt = '0.00%';  
           }else{
          worksheet.getCell(this.arr_colume[ax].name_col+"154").value = this.ac_width_kpi/100
          worksheet.getCell(this.arr_colume[ax].name_col+"154").numFmt = '0.00%';
           }
          
          this.acc_width_arr.push({
              value:this.ac_width_kpi
          })
          
        }

        
         for(var ax = 0; ax<this.sum_ttl_marker.length; ax++){
         worksheet.getCell(this.arr_colume[ax].name_col+"155").value = Number(this.count_work_week[ax].p)
         
         }
        
        for(var ax = 0; ax<this.sum_ttl_marker.length; ax++){
        this.avg_end_width = 0
        this.finish_avg_end_width = 0
        if(isNaN(this.end_loss_target_arr[ax].value)==true || isFinite(this.end_loss_target_arr[ax].value)==false){
          this.end_loss_target_arr[ax].value = 0
        }
        if(isNaN(this.acc_width_arr[ax].value)==true || isFinite(this.acc_width_arr[ax].value)==false){
          this.acc_width_arr[ax].value = 0
        }
        this.avg_end_width = (Number(this.end_loss_target_arr[ax].value)+Number(this.acc_width_arr[ax].value))/(2)
    
        if(isNaN(this.avg_end_width)==true || isFinite(this.avg_end_width)==false || this.avg_end_width <= 0){
          worksheet.getCell(this.arr_colume[ax].name_col+"156").value = 0.00/100
          worksheet.getCell(this.arr_colume[ax].name_col+"156").numFmt = '0.00%';  
           }else{
          worksheet.getCell(this.arr_colume[ax].name_col+"156").value = this.avg_end_width/100
          worksheet.getCell(this.arr_colume[ax].name_col+"156").numFmt = '0.00%';
           }
        }
      worksheet.getCell("X157").value =11
      worksheet.mergeCells("X157:AA157")

      worksheet.getCell("AB157").value =12
      worksheet.mergeCells("AB157:AF157")

      worksheet.getCell("AG157").value =1
      worksheet.mergeCells("AG157:AJ157")

      worksheet.getCell("AK157").value =2
      worksheet.mergeCells("AK157:AN157")

      worksheet.getCell("AO157").value =3
      worksheet.mergeCells("AO157:AS157")

      worksheet.getCell("AT157").value =4
      worksheet.mergeCells("AT157:AW157")

      worksheet.getCell("AX157").value =5
      worksheet.mergeCells("AX157:BA157")

      worksheet.getCell("BB157").value =6
      worksheet.mergeCells("BB157:BF157")

      worksheet.getCell("BG157").value =7
      worksheet.mergeCells("BG157:BJ157")

      worksheet.getCell("BK157").value =8
      worksheet.mergeCells("BK157:BO157")

      worksheet.getCell("BP157").value =9
      worksheet.mergeCells("BP157:BS157")

      worksheet.getCell("BT157").value =10
      worksheet.mergeCells("BT157:BW157")

  
  
    this.week_of_set = ""
    for(var ax = 0; ax<this.date_use_data.length; ax++){
      if(this.date_use_data[ax].convert_week == this.start_date && this.date_use_data[ax].convert_week2 == this.end_date){
        this.week_of_set = this.date_use_data[ax].value
      }
    }
 worksheet.getCell("A1").value = "Spread Loss Audit By Fabric Type Weekly Summary NY"+this.org+" ( MU-SLA-3-2 ) : Period  : "+ this.week_of_set +"   : "+this.start+"-"+this.end; 
       worksheet.mergeCells("A1:S2");
       worksheet.getCell("A1").font = {
        name: 'Angsana New',
        color: { argb: 'FF3333FF' },
        family: 4,
        size: 30,
        bold: true,
        italic:true
      };
        worksheet.getCell("A1").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
       worksheet.getCell("A1").alignment = {
          horizontal: "left",
          vertical: "middle",
        };
       
      this.$q.loading.hide({
          
        })

        worksheet.mergeCells('A15:Q75')
        worksheet.getCell('A15').fill={
          type: 'pattern',
          pattern:'none',
          fgColor:{argb:'FFFFFFFF'},
          bgColor:{argb:'FFFFFFFF'},
        };
            workbook.xlsx.writeBuffer().then((data) => {
        const blob = new Blob([data], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8",
        });
        saveAs(blob, "Report Weekly SLA3-2.xlsx");
      });    
      
       for(var ax = 0; ax<this.total_spread_loss_woven_data.length; ax++){
        if(this.total_spread_loss_woven_data[ax].value > 0 ){
          this.total_spread_loss_woven_data_find_0.push({
            value:this.total_spread_loss_woven_data[ax].value,
            value_ax:ax
          })
        }
      }
     
      for(var ax = 0; ax<this.total_spread_loss_woven_data_find_0.length; ax++){
         this.date_data_for_graph1.push({
           value:this.date_use_data[this.total_spread_loss_woven_data_find_0[ax].value_ax].first_date+" - "+ this.date_use_data[this.total_spread_loss_woven_data_find_0[ax].value_ax].sec_date
         })

         this.data_graph1.push({
           value:this.total_spread_loss_woven_data[this.total_spread_loss_woven_data_find_0[ax].value_ax].value
         })

         this.data_target_graph1.push({
           value:this.total_spread_loss_woven_target[this.total_spread_loss_woven_data_find_0[ax].value_ax].value
         }) 

        
      } 
     
        
           let woven_data = [];
            let index1 = 0;
             this.total_spread_loss_woven_data .forEach(e => {
               woven_data[index1] = parseFloat(e.value).toFixed(2)
                index1++
            })

            let woven_target = [];
            let index2 = 0;
             this.total_spread_loss_woven_target .forEach(e => {
                woven_target[index2] = e.value
                index2++
            })
         
            let data_date =  [];
            let index = 0;
           this.date_data_for_graph1 .forEach(e => {
                data_date[index] = e.value
                index++
            })
        
   


var myChart = echarts.init(chartDom);


var option;


option = {
  title: 
    {
      left: 'center',
      text: 'Spread Loss Audit Weekly Summary'
    },
    tooltip: {
    trigger: 'axis',
    axisPointer: {
      animation: false
    }
    },

    legend: {
    left: 'left',
    data: ['actual', 'target'],
  },
  xAxis: {
    type: 'category',
    data: data_date
  },
  yAxis: {
    type: 'value',
    boundaryGap: [0, '200%']
  },
  series: [
    {
      name:'actual',
      data: woven_data,
      type: 'line'
    },
    {
      name:'target',
      data: woven_target,
      type: 'line'
    }
  ]
};

option && myChart.setOption(option);



 for(var ax = 0; ax<this.total_spread_loss_light_data.length; ax++){
        if(this.total_spread_loss_light_data[ax].value > 0 ){
          this.total_spread_loss_light_data_find_0.push({
            value:this.total_spread_loss_light_data[ax].value,
            value_ax:ax
          })
        }
      }
 
      for(var ax = 0; ax<this.total_spread_loss_light_data_find_0.length; ax++){
         this.date_data_for_graph2.push({
           value:this.date_use_data[this.total_spread_loss_light_data_find_0[ax].value_ax].first_date+" - "+ this.date_use_data[this.total_spread_loss_light_data_find_0[ax].value_ax].sec_date
         })

         this.data_graph2.push({
           value:this.total_spread_loss_light_data[this.total_spread_loss_light_data_find_0[ax].value_ax].value
         })

         this.data_target_graph2.push({
           value:this.total_spread_loss_light_target[this.total_spread_loss_light_data_find_0[ax].value_ax].value
         }) 

        
      } 

            let light_data = [];
            let index3 = 0;
             this.data_graph2 .forEach(e => {
                light_data[index3] = e.value
                index3++
            })

            let light_target = [];
            let index4 = 0;
            this.data_target_graph2 .forEach(e => {
                light_target[index4] = e.value
                index4++
            })

            let data_date2 = [];
            let index10 = 0;
          this.date_data_for_graph2 .forEach(e => {
                data_date2[index10] = e.value
                index10++
            })



var myChart2 = echarts.init(chartDom2);
var option2;


option2 = {
  
  title: 
    {
      left: 'center',
      text: 'Spread Loss Audit Weekly Summary'
    },
    tooltip: {
    trigger: 'axis',
    axisPointer: {
      animation: false
    }
    },

    legend: {
    left: 'left',
    data: ['actual', 'target'],
  },
  xAxis: {
    type: 'category',
    data: data_date2
  },
  yAxis: {
    type: 'value',
     boundaryGap: [0, '30%']
  },
  series: [
    {
      name: 'actual',
      data: light_data,
      type: 'line'
    },
    {
      name: 'target',
      data: light_target,
      type: 'line'
    }
  ]
};

option2 && myChart2.setOption(option2);

for(var ax = 0; ax<this.total_spread_loss_fleece_data.length; ax++){
        if(this.total_spread_loss_fleece_data[ax].value > 0 ){
          this.total_spread_loss_fleece_data_find_0.push({
            value:this.total_spread_loss_fleece_data[ax].value,
            value_ax:ax
          })
        }
      }
 
      for(var ax = 0; ax<this.total_spread_loss_fleece_data_find_0.length; ax++){
         this.date_data_for_graph3.push({
           value:this.date_use_data[this.total_spread_loss_fleece_data_find_0[ax].value_ax].first_date+" - "+ this.date_use_data[this.total_spread_loss_fleece_data_find_0[ax].value_ax].sec_date
         })

         this.data_graph3.push({
           value:this.total_spread_loss_fleece_data[this.total_spread_loss_fleece_data_find_0[ax].value_ax].value
         })

         this.data_target_graph3.push({
           value:this.total_spread_loss_fleece_target[this.total_spread_loss_fleece_data_find_0[ax].value_ax].value
         }) 

        
      } 


            let fleece_data = [];
            let index5 = 0;
          this.data_graph3 .forEach(e => {
                fleece_data[index5] = e.value
                index5++
            })

            let fleece_target = [];
            let index6 = 0;
            this.data_target_graph3 .forEach(e => {
                fleece_target[index6] = e.value
                index6++
            })

            let date_data3 = [];
            let index11 = 0;
             this.date_data_for_graph3 .forEach(e => {
                date_data3[index11] = e.value
                index11++
            })


var myChart3 = echarts.init(chartDom3);
var option3;


option3 = {
  title: 
    {
      left: 'center',
      text: 'Spread Loss Audit Weekly Summary'
    },
    tooltip: {
    trigger: 'axis',
    axisPointer: {
      animation: false
    }
    },

    legend: {
    left: 'left',
    data: ['actual', 'target'],
  },
  xAxis: {
    type: 'category',
    data: date_data3
  },
  yAxis: {
    type: 'value',
    boundaryGap: [0, '30%']
  },
  series: [
    {
      name: 'actual',
      data: fleece_data,
      type: 'line'
    },
    {
      name: 'target',
      data: fleece_target,
      type: 'line'
    }
  ]
};

option3 && myChart3.setOption(option3);


for(var ax = 0; ax<this.total_spread_loss_special_data.length; ax++){
        if(this.total_spread_loss_special_data[ax].value > 0 ){
          this.total_spread_loss_special_data_find_0.push({
            value:this.total_spread_loss_special_data[ax].value,
            value_ax:ax
          })
        }
      }
 
      for(var ax = 0; ax<this.total_spread_loss_special_data_find_0.length; ax++){
         this.date_data_for_graph4.push({
           value:this.date_use_data[this.total_spread_loss_special_data_find_0[ax].value_ax].first_date+" - "+ this.date_use_data[this.total_spread_loss_special_data_find_0[ax].value_ax].sec_date
         })

         this.data_graph4.push({
           value:this.total_spread_loss_special_data[this.total_spread_loss_special_data_find_0[ax].value_ax].value
         })

         this.data_target_graph4.push({
           value:this.total_spread_loss_special_target[this.total_spread_loss_special_data_find_0[ax].value_ax].value
         }) 

        
      } 

            let special_data = [];
            let index7 = 0;
             this.data_graph4 .forEach(e => {
                special_data[index7] = e.value
                index7++
            })

            let special_target = [];
            let index8 = 0;
            this.data_target_graph4 .forEach(e => {
                special_target[index8] = e.value
                index8++
            })

            let data_date4 = [];
            let index12 = 0;
            this.date_data_for_graph4 .forEach(e => {
                data_date4[index12] = e.value
                index12++
            })



var myChart4 = echarts.init(chartDom4);
var option4;


option4 = {
  title: 
    {
      left: 'center',
      text: 'Spread Loss Audit Weekly Summary'
    },
    tooltip: {
    trigger: 'axis',
    axisPointer: {
      animation: false
    }
    },

    legend: {
    left: 'left',
    data: ['actual', 'target'],
  },
  xAxis: {
    type: 'category',
    data: data_date4
  },
  yAxis: {
    type: 'value',
    boundaryGap: [0, '10%']
  },
  series: [
    {
      name: 'actual',
      data: special_data,
      type: 'line'
    },
    {
      name: 'target',
      data: special_target,
      type: 'line'
    }
  ]
};

option4 && myChart4.setOption(option4);

    }catch(err){
      this.$q.loading.hide({
          
        })
       this.$q.notify({
          message: "Something is Wrong",
          color: "red-9",
        });
        console.log(err)
    }
    
},  
},
 watch:{
        start(val){
          let day = val.substring(0, 2);
          let month = val.substring(3, 5);
          let year = val.substring(6, 10);
          this.start_date = year +"/"+ month +"/"+ day

        },
        end(val){
           let day = val.substring(0, 2);
          let month = val.substring(3, 5);
          let year = val.substring(6, 10);
          this.end_date = year +"/"+ month +"/"+ day

        } 
  },
}
</script>

<style>
.myclass {
  padding-top : 50px;
  padding-right: 40px;
  padding-bottom:40px;
  padding-left:50px;
}
</style>