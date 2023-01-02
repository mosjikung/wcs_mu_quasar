<template>
<div class="bg-blue-10 text-white">
      <q-toolbar>
        <q-btn flat round dense />

        <q-toolbar-title>Weekly Mu SLA 3-1</q-toolbar-title>

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
  
    <q-select
      class="q-pa-xl"
      v-model="org"
      :options="option_org"
      style="width: 200px"
      disable
      label="Chose Org"
    />

    <div class="myclass">  <q-btn
      size="lg"
      dense
      color="primary"
      label="Export"
      @click="exportexcel()"
    /></div>
    <div class="q-pa-xl" style="min-width:100%;" >
  <q-banner inline-actions rounded class="bg-orange-4 text-white" >
    <h5 class="q-mt-xs" >Graph For Mu Sla 3-1</h5>
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
    <!-- 
     <q-btn
      size="md"
      dense
      class="q-px-xl q-py-xl"
      color="primary"
      label="Update"
      @click="UpdateData()"
    /> -->
   
    
     
  </div>
</template>

<script>
import { ref } from "vue";
import { date } from "quasar";
import axios from "axios";
import { useQuasar, QSpinnerGears } from 'quasar'
import { onBeforeUnmount } from 'vue'
import * as Excel from "exceljs";
import { saveAs } from "file-saver";
import moment from "moment";
import * as echarts from 'echarts';

export default {
  data(){ //data
      return{ 
        start:"",
        end:"",
        org:"",
        date_use_data:[],
        date_use_data2:[],
        date_use_data_re:[],
        ttl_marker_yd:[],
        ttl_marker_yd2:[],
        ac_end_loss2:[],
        ac_splice2:[],
        ac_cut2:[],
        ac_rem2:[],
        keep_history_data:[],
        all_data_arr:[],
        all_data_arr2:[],
        end_week:[],
        end_week_change:[],
        splice_week:[],
        splice_week_change:[],
        cut_week:[],
        cut_week_change:[],
        rem_week:[],
        rem_week_change:[],
        val_end:[],
        val_splice:[],
        val_cut:[],
        val_rem:[],
        width_loss_yd:[],
        width_loss_yd_change:[],
        width_loss_yd_t:[],
        ttl_marker_yd_t:[],
        all_week_cut:[],
        all_week_cut_2:[],
        all_week_cut_4:[],
        all_week_cut4_4:[],
        all_week_cut4_4_change:[],
        all_week_yard:[],
        commulate_spread_yd:[],
        rowexport:[],
        ttl_marker_1_4:[],
        ttl_marker_1_4_all:[],
        ttl_obswidth:[],
        total_spread_width_loss:[],
        keep_ac_width:[],
        all_ttl_marker1_4_diff:[],
        result_diff:[],
        result_actual:[],
        rowx:[],
        all_week_cut_yard_4_4_1:[],
        all_week_cut_yard_4_4_1_change:[],
        all_week_cut_yard_4_4_2:[],
        all_week_cut_yard_4_4_2_change:[],
        ttl_spread_graph:[],
        ttl_spread_graph_last:[],
        
        ttl_target:[],
        date_data_for_graph:[],

        avg_width_loss_cal:[],
        status_week_open:true,

        ttl_marker_yd_change:[],
        width_loss_yd_t_change:[],

        ttl_marker_endless_woven: [],
        ttl_marker_endless_light: [],
        ttl_marker_endless_fleece: [],
        ttl_marker_endless_special: [],

        ttl_marker_splice_woven: [],
        ttl_marker_splice_light: [],
        ttl_marker_splice_fleece: [],
        ttl_marker_splice_special: [],

        ttl_marker_cut_woven: [],
        ttl_marker_cut_light: [],
        ttl_marker_cut_fleece: [],
        ttl_marker_cut_special: [],

        ttl_marker_rem_woven: [],
        ttl_marker_rem_light: [],
        ttl_marker_rem_fleece: [],
        ttl_marker_rem_special: [],


        total_target_non_0:[],
        
        all_minus_value:[],
        
        


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

   

        year:[],
        end_loss_value:[],
        splice_loss_value:[],
        cut_out_loss_value:[],
        remnants_loss_value:[],
        total_spread_loss_per_value:[],
        total_width_loss_per_value:[],


        total_target:[],
        
        

        keep_end_woven:[],
        keep_end_light:[],
        keep_end_fleece:[],
        keep_end_special:[],

        keep_splice_woven:[],
        keep_splice_light:[],
        keep_splice_fleece:[],
        keep_splice_special:[],

        keep_cut_woven:[],
        keep_cut_light:[],
        keep_cut_fleece:[],
        keep_cut_special:[],

        keep_rem_woven:[],
        keep_rem_light:[],
        keep_rem_fleece:[],
        keep_rem_special:[],


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
        total_loss:[
          {
          p:0.38
          },
           {
          p:0.50
          },
           {
          p:0.56
          },
           {
          p:0.65
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

        column_down:[
          {
            cn:"T"
          },
           {
            cn:"U"
          },
           {
            cn:"V"
          },
           {
            cn:"W"
          },
           {
            cn:"X"
          },
        ],
        colume_arr:[
          {
            cn:"H"
          },
           {
            cn:"I"
          },
           {
            cn:"J"
          },
           {
            cn:"K"
          },
        ],
         arr_colume2:[
        {
          name_col:"T"
        },
        {
          name_col:"U"
        },
        {
          name_col:"V"
        },
         {
          name_col:"W"
        },
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
       
       
        
      ],
      
      arr_column3:[
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
        
        first_day:"2021/10/25",
        last_day:"2022/10/24",
      
      }
  },
 async mounted() {
    this.year=[]
      this.end_loss_value=[]
      this.splice_loss_value=[]
      this.cut_out_loss_value=[]
      this.remnants_loss_value=[]
      this.total_spread_loss_per_value=[]
      this.total_width_loss_per_value=[]
    this.rowx = []
    let org = this.$q.localStorage.getItem("org");
    this.org = org

    await axios.get(this.$api_url + "/mainso.php/get_target")
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
     await axios.get(this.$api_url + "/mainso.php/get_date2")
      .then((resp)=>{
        this.first_day = (resp.data.data[0].START_DATE)
      }).catch((error)=>{
        console.log(error)
      })

      const params2 = new FormData();
      params2.append("org",this.org)
     await axios({
        method:'post',
        url: this.$api_url + '/mainso.php/find_history_data',
        data:params2,
      }).then((resp)=>{
        console.log(resp.data.data)
        this.keep_history_data = resp.data.data
      })
      
      
      for(var ax = 0; ax<this.keep_history_data.length; ax++){
        this.year[ax] = this.keep_history_data[ax].YEAR
        this.end_loss_value[ax] = this.keep_history_data[ax].END_LOSS
        this.splice_loss_value[ax] = this.keep_history_data[ax].SPLICE_LOSS
        this.cut_out_loss_value[ax] = this.keep_history_data[ax].CUT_OUT_LOSS
        this.remnants_loss_value[ax] = this.keep_history_data[ax].REMNANTS_LOSS
        this.total_spread_loss_per_value[ax] = this.keep_history_data[ax].TOTAL_SPREAD_LOSS_PER
        this.total_width_loss_per_value[ax] = this.keep_history_data[ax].TOTAL_WIDTH_LOSS_PER
      }

   

    


  },
methods: {
    async exportexcel() {
   /*    try{ */
      this.$q.loading.show({
          spinner: QSpinnerGears,
          spinnerColor: 'wthite',
          spinnerSize: 180,
          backgroundColor: 'black',
         
        })
      this.date_use_data=[]
      this.date_use_date2=[]
      this.date_use_data2 = []
      this.date_use_data2 =[]
      this.data_use_data_re =[]
      this.ttl_marker_yd = []
      this.ttl_marker_yd2 = []
      this.ac_end_loss2=[]
      this.ac_splice2 =[]
      this.ac_cut2 =[]
      this.ac_rem2=[]
      this.end_week=[]
      this.end_week_change=[]
      this.splice_week =[]
      this.splice_week_change = []
      this.cut_week=[]
      this.cut_week_change=[]
      this.rem_week=[]
      this.rem_week_change = []
      this.val_end=[]
      this.val_splice=[]
      this.val_cut=[]
      this.val_rem=[]
      this.width_loss_yd=[]
      this.width_loss_yd_change=[]
      this.width_loss_yd_t=[]
      this.width_loss_yd_t_change=[]
      this.ttl_marker_yd_t=[]
      this.all_week_yard=[]
      this.all_week_cut=[]
      this.all_week_cut_2=[]
      this.all_week_cut_4=[]
      this.all_week_cut4_4=[]
      this.all_week_cut4_4_change=[]
      this.ttl_marker_1_4=[]
      this.commulate_spread_yd=[]
      this.ttl_marker_1_4_all=[]
      this.ttl_obswidth=[]  //array[]
      this.keep_ac_width=[]
      this.all_ttl_marker1_4_diff=[]
      this.total_spread_width_loss=[]
      this.result_diff=[]
      this.result_actual=[],
      this.all_week_cut_yard_4_4_1=[]
      this.all_week_cut_yard_4_4_1_change=[]
      this.all_week_cut_yard_4_4_2=[]
      this.all_week_cut_yard_4_4_2_change=[]
      this.ttl_spread_graph=[]
      this.total_target=[]
      this.total_target_non_0=[]
      this.date_data_for_graph=[]
      this.ttl_spread_graph_last=[]
      this.avg_width_loss_cal = []
      this.all_minus_value =[]

        this.keep_end_woven=[]
        this.keep_end_light=[]
        this.keep_end_fleece=[]
        this.keep_end_special=[]

        this.keep_splice_woven=[]
        this.keep_splice_light=[]  //ตรงนี้
        this.keep_splice_fleece=[]
        this.keep_splice_special=[]

        this.keep_cut_woven=[]
        this.keep_cut_light=[]
        this.keep_cut_fleece=[]
        this.keep_cut_special=[]

        this.ttl_marker_yd_change=[],

        this.keep_rem_woven=[]
        this.keep_rem_light=[]
        this.keep_rem_fleece=[]
        this.keep_rem_special=[]

        this.ttl_marker_endless_woven = []
        this.ttl_marker_endless_light = []
        this.ttl_marker_endless_fleece = []
        this.ttl_marker_endless_special = []

        this.ttl_marker_splice_woven= []
        this.ttl_marker_splice_light= []
        this.ttl_marker_splice_fleece= []
        this.ttl_marker_splice_special= []

        this.ttl_marker_cut_woven= []
        this.ttl_marker_cut_light= []
        this.ttl_marker_cut_fleece= []
        this.ttl_marker_cut_special= []

        this.ttl_marker_rem_woven= []
        this.ttl_marker_rem_light= []
        this.ttl_marker_rem_fleece= []
        this.ttl_marker_rem_special= []

      

        this.ttl_target=[]
        
        


        const workbook = new Excel.Workbook();
      workbook.creator = "Nanyang";
      workbook.lastModifiedBy = "Admin";
      workbook.created = new Date(2021, 8, 30);
      workbook.modified = new Date();
      workbook.lastPrinted = new Date(2021, 7, 27);
      const worksheet = workbook.addWorksheet("Data");

      worksheet.columns = [
        { key: "A", width: 18 },
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
        { key: "N", width: 15},
        { key: "O", width: 15},
        { key: "P", width: 15 },
        { key: "Q", width: 6.7 },
        { key: "R", width: 18 },
        { key: "S", width: 1 },
        { key: "T", width: 13 },
        { key: "U", width: 6.7 },
        { key: "V", width: 6.7 },
        { key: "W", width: 13 },
        { key: "X", width: 6.7 },
        { key: "AC", width: 15.13 },
      ];

     
     

        worksheet.getCell("A3").value = "Descriptions"; 
       worksheet.mergeCells("A3:A5");
   

          worksheet.getCell("B3").value = "History Performance"; 
       worksheet.mergeCells("B3:F3");
   

          worksheet.getCell("B4").value = Number(this.year[0]);
          
       worksheet.mergeCells("B4:B5");
    

          worksheet.getCell("C4").value = Number(this.year[1]); 
       worksheet.mergeCells("C4:C5");


      worksheet.getCell("D4").value = Number(this.year[2]); 
       worksheet.mergeCells("D4:D5");
 


          worksheet.getCell("E4").value = Number(this.year[3]); 
       worksheet.mergeCells("E4:E5");
     


          worksheet.getCell("F4").value = Number(this.year[4]); 
       worksheet.mergeCells("F4:F5");
  

        
       


           worksheet.getCell("G3").value = "Weighted Target of This Week"; 
       worksheet.mergeCells("G3:G4");
     
       worksheet.getCell("G5").value = "Overall"; 
      
  
        //Actual Performance
            worksheet.getCell("H3").value = "Actual Performance"; 
       worksheet.mergeCells("H3:M3");


        
 

             worksheet.getCell("L4").value = "Cumulated 4 weeks "; 
       worksheet.mergeCells("L4:M4");
   

    

           worksheet.getCell("L5").value = "Average"; 


           worksheet.getCell("M5").value = "Total"; 


      worksheet.getCell("N3").value = "Cumulated Since"; 
       worksheet.mergeCells("N3:O4");
 


           worksheet.getCell("N5").value = "Average"; 
   

           worksheet.getCell("O5").value = "Total"; 

           worksheet.getCell("A6").value = "Total  Marker  (Yds)"; 
           worksheet.getCell("A7").value = "End Loss";
           worksheet.getCell("A8").value = "Splice Loss"; 
           worksheet.getCell("A9").value = "Cut out Loss"; 
           worksheet.getCell("A10").value = "Remnants Loss";
           worksheet.getCell("A11").value = "% Total  Spread  Loss";
           worksheet.getCell("A12").value = "Total  Spread  Loss  (Yds)";
           worksheet.getCell("A13").value = "Total Width Loss (Yds)";
           worksheet.getCell("A14").value = "Total Width Loss (%)";
           worksheet.getCell("A15").value = "Total Fabric Cut (Yds)";
           worksheet.getCell("A16").value = "Audit per Total Fabric Cut";
           worksheet.getCell("A17").value = "Total Dozens Cut";

      
      
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

        
        worksheet.getCell(this.column_history[axc].name_col+14).value = this.total_width_loss_per_value[axc]/100
        worksheet.getCell(this.column_history[axc].name_col+14).numFmt = "0.00%"

      }
         
         
            
      for(var ax = 0; ax<this.arr_column3.length; ax++){
        for(var bz = 1; bz<18;  bz++){
        worksheet.getCell(this.arr_column3[ax].name_col+bz).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
        worksheet.getCell(this.arr_column3[ax].name_col+bz).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
       worksheet.getCell(this.arr_column3[ax].name_col+bz).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
      }
      }

      for(var ax = 6; ax<17; ax++){
        worksheet.getCell("R"+ax).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
       
       worksheet.getCell("R"+ax).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
      }

      for(var ax = 18; ax<22; ax++){
        worksheet.getCell("R"+ax).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
      
       worksheet.getCell("R"+ax).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
      }

      for(var ax = 0; ax<this.arr_colume2.length; ax++){
        for(var bz = 4; bz<16;  bz++){
        worksheet.getCell(this.arr_colume2[ax].name_col+bz).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
       
       worksheet.getCell(this.arr_colume2[ax].name_col+bz).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
      }
      }

      for(var ax = 0; ax<this.arr_colume2.length; ax++){
        for(var bz = 18; bz<22;  bz++){
        worksheet.getCell(this.arr_colume2[ax].name_col+bz).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
       
       worksheet.getCell(this.arr_colume2[ax].name_col+bz).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
      }
      }

      for(var ax = 0; ax<this.arr_colume2.length; ax++){
        for(var bz = 18; bz<22;  bz++){
        worksheet.getCell(this.arr_colume2[ax].name_col+bz).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
      
       worksheet.getCell(this.arr_colume2[ax].name_col+bz).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
      }
      }

      for(var ax = 0; ax<this.column_down.length-1; ax++){
        for(var bz = 24; bz<29;  bz++){
        worksheet.getCell(this.column_down[ax].cn+bz).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
    
       worksheet.getCell(this.column_down[ax].cn+bz).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
      }
      }

      for(var ax = 0; ax<this.column_down.length; ax++){
        for(var bz = 32; bz<39;  bz++){
        worksheet.getCell(this.column_down[ax].cn+bz).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
      
       worksheet.getCell(this.column_down[ax].cn+bz).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
      }
      }
         


        

      const params5 = new FormData();
      params5.append("start",this.start_date)
      params5.append("end",this.end_date)
      let org = this.$q.localStorage.getItem("org");
      params5.append("org",org)
      for (var pair of params5.entries()) {
                  console.log(pair[0] + ", " + pair[1]);
                }

       await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data_weekly',
        data:params5,
      }).then((resp)=>{
        
        
        if (resp.data.data.length > 0) {
           this.rowexport=[]
          resp.data.data.forEach((e) => {
            this.rowexport.push(
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
              ttl_marker: e.TTL_MARKER_YD,
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
              balance: 
                e.ISSUE - (e.TTL_MARKER_YD * e.WEIGHT_YD) / 1000
              ,
            })
          
          });
             this.all_result_ttlmarker_woven2 = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "1"){
              this.all_result_ttlmarker_woven2 = Number(this.all_result_ttlmarker_woven2) + Number(this.rowexport[ax].ttl_marker)
            }
          }
          this.ttl_marker_1_4.push({
            value:this.all_result_ttlmarker_woven2
          })

          this.all_result_ttlmarker_light2 = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "2"){
              this.all_result_ttlmarker_light2 = Number(this.all_result_ttlmarker_light2) + Number(this.rowexport[ax].ttl_marker)
            }
          }
          this.ttl_marker_1_4.push({
            value:this.all_result_ttlmarker_light2
          })

          this.all_result_ttlmarker_fleece2 = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "3"){
              this.all_result_ttlmarker_fleece2 = Number(this.all_result_ttlmarker_fleece2) + Number(this.rowexport[ax].ttl_marker)
            }
          }
          this.ttl_marker_1_4.push({
            value:this.all_result_ttlmarker_fleece2
          })

          this.all_result_ttlmarker_special2 = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "4"){
              this.all_result_ttlmarker_special2 = Number(this.all_result_ttlmarker_special2) + Number(this.rowexport[ax].ttl_marker)
            }
          }
          this.ttl_marker_1_4.push({
            value:this.all_result_ttlmarker_special2
          })

          this.all_result_ttlmarker = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
           
              this.all_result_ttlmarker = Number(this.all_result_ttlmarker) + Number(this.rowexport[ax].ttl_marker)
            
          }
          this.ttl_marker_1_4_all.push({
            value:this.all_result_ttlmarker
          })

          

          this.all_result_ttlmarker_woven = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "1"){
              this.all_result_ttlmarker_woven = Number(this.all_result_ttlmarker_woven) + Number(this.rowexport[ax].ttl_marker)
            }
          }
         
           this.all_ttl_marker1_4_diff.push({
            value_ttl : this.all_result_ttlmarker_woven
          })

      this.all_result_ttlmarker_light = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "2"){
              this.all_result_ttlmarker_light = Number(this.all_result_ttlmarker_light) + Number(this.rowexport[ax].ttl_marker)
            }
          }
       
           this.all_ttl_marker1_4_diff.push({
            value_ttl : this.all_result_ttlmarker_light
          })
          

      this.all_result_ttlmarker_fleece = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "3"){
              this.all_result_ttlmarker_fleece = Number(this.all_result_ttlmarker_fleece) + Number(this.rowexport[ax].ttl_marker)
            }
          }
           this.all_ttl_marker1_4_diff.push({
            value_ttl : this.all_result_ttlmarker_fleece
          })

         
      this.all_result_ttlmarker_special = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "4"){ 
              this.all_result_ttlmarker_special = Number(this.all_result_ttlmarker_special) + Number(this.rowexport[ax].ttl_marker)
            }
          }
          
           this.all_ttl_marker1_4_diff.push({
            value_ttl : this.all_result_ttlmarker_special
          })


          this.all_val_end_woven = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "1"){
              this.all_val_end_woven = Number(this.all_val_end_woven) + Number(this.rowexport[ax].endless_length_yd)
            }
          }
          this.keep_end_woven.push({
            value : this.all_val_end_woven
          })

          this.all_val_end_light = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "2"){
              this.all_val_end_light = Number(this.all_val_end_light) + Number(this.rowexport[ax].endless_length_yd)
            }
          }
          this.keep_end_light.push({
            value :this.all_val_end_light
          })

          this.all_val_end_fleece = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "3"){
              this.all_val_end_fleece = Number(this.all_val_end_fleece) + Number(this.rowexport[ax].endless_length_yd)
            }
          }
          this.keep_end_fleece.push({
            value : this.all_val_end_fleece
          })

          this.all_val_end_special = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "4"){
              this.all_val_end_special = Number(this.all_val_end_special) + Number(this.rowexport[ax].endless_length_yd)
            }
          }
          this.keep_end_special.push({
            value : this.all_val_end_special
          })

        //end

        this.all_val_splice_woven = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "1"){
              this.all_val_splice_woven = Number(this.all_val_splice_woven) + Number(this.rowexport[ax].spliceloss)
            }
          }
          this.keep_splice_woven.push({
            value : this.all_val_splice_woven
          })

          this.all_val_splice_light = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            
            if(this.rowexport[ax].f_type == "2"){
              this.all_val_splice_light = Number(this.all_val_splice_light) + Number(this.rowexport[ax].spliceloss)
              
            }
          }
          this.keep_splice_light.push({
            value :this.all_val_splice_light
          })
          

          this.all_val_splice_fleece = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "3"){
              this.all_val_splice_fleece = Number(this.all_val_splice_fleece) + Number(this.rowexport[ax].spliceloss)
            }
          }
          this.keep_splice_fleece.push({
            value : this.all_val_splice_fleece
          })

          this.all_val_splice_special = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "4"){
              this.all_val_splice_special = Number(this.all_val_splice_special) + Number(this.rowexport[ax].spliceloss)
            }
          }
          this.keep_splice_special.push({
            value : this.all_val_splice_special
          })

          //splice


          this.all_val_cut_woven = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "1"){
              this.all_val_cut_woven = Number(this.all_val_cut_woven) + Number(this.rowexport[ax].totalcutout)
            }
          }
          this.keep_cut_woven.push({
            value : this.all_val_cut_woven
          })

          this.all_val_cut_light = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "2"){
            
              this.all_val_cut_light = Number(this.all_val_cut_light) + Number(this.rowexport[ax].totalcutout)
            }
          }
          this.keep_cut_light.push({
            value :this.all_val_cut_light
          })

          this.all_val_cut_fleece = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "3"){
              this.all_val_cut_fleece = Number(this.all_val_cut_fleece) + Number(this.rowexport[ax].totalcutout)
            }
          }
          this.keep_cut_fleece.push({
            value : this.all_val_cut_fleece
          })

          this.all_val_cut_special = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "4"){
              this.all_val_cut_special = Number(this.all_val_cut_special) + Number(this.rowexport[ax].totalcutout)
            }
          }
          this.keep_cut_special.push({
            value : this.all_val_cut_special
          })

          //cut


            this.all_val_rem_woven = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "1"){
              this.all_val_rem_woven = Number(this.all_val_rem_woven) + Number(this.rowexport[ax].rement)
            }
          }
          this.keep_rem_woven.push({
            value : this.all_val_rem_woven
          })

          this.all_val_rem_light = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "2"){
             
              this.all_val_rem_light = Number(this.all_val_rem_light) + Number(this.rowexport[ax].rement)
              
            }
          }
          this.keep_rem_light.push({
            value :this.all_val_rem_light
          })

          this.all_val_rem_fleece = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "3"){
              this.all_val_rem_fleece = Number(this.all_val_rem_fleece) + Number(this.rowexport[ax].rement)
            }
          }
          this.keep_rem_fleece.push({
            value : this.all_val_rem_fleece
          })

          this.all_val_rem_special = 0
          for (var ax = 0; ax<this.rowexport.length; ax++){
            if(this.rowexport[ax].f_type == "4"){
              this.all_val_rem_special = Number(this.all_val_rem_special) + Number(this.rowexport[ax].rement)
            }
          }
          this.keep_rem_special.push({
            value : this.all_val_rem_special
          })

        //rem


        
        }else{
          this.rowexport.push({
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
              this.all_ttl_marker1_4_woven.push({
            value_ttl : 0
          })
             this.all_ttl_marker1_4_light.push({
            value_ttl : 0
          })
             this.all_ttl_marker1_4_fleece.push({
            value_ttl : 0
          })
             this.all_ttl_marker1_4_special.push({
            value_ttl : 0
          })
          this.keep_end_woven.push({
            value : 0
          })
          this.keep_end_light.push({
            value : 0
          })
          this.keep_end_fleece.push({
            value : 0
          })
          this.keep_end_special.push({
            value : 0
          })

           this.keep_splice_woven.push({
            value : 0
          })
          this.keep_splice_light.push({
            value : 0
          })
          this.keep_splice_fleece.push({
            value : 0
          })
          this.keep_splice_special.push({
            value : 0
          })

           this.keep_cut_woven.push({
            value : 0
          })
          this.keep_cut_light.push({
            value : 0
          })
          this.keep_cut_fleece.push({
            value : 0
          })
          this.keep_cut_special.push({
            value : 0
          })

           this.keep_rem_woven.push({
            value : 0
          })
          this.keep_rem_light.push({
            value : 0
          })
          this.keep_rem_fleece.push({
            value : 0
          })
          this.keep_rem_special.push({
            value : 0
          })
        }
        
      
      })
      .catch((error)=>{
         this.$q.notify({
          message: "Something is error Please Try again",
          timeout: 2000,
          position:"center",
          color: "red-9",
        });

        
        this.$q.loading.hide({
          
        })
        
        
        console.log(error)
      })
      
     if(this.rowexport.length >1){
      
      var myNewDate = new Date(this.start_date);
        var myNewDate2 = new Date(myNewDate);
         var count = 0;
        
      for(var ax = 0; ax<4; ax++){
       if(count == 0 ){
          this.firstNewDate = moment(myNewDate).format("DD/MM");
          this.convert_axios1 = moment(myNewDate).format("YYYY-MM-DD");
          

          this.next_date =  myNewDate.setDate(myNewDate.getDate() + 5);
          this.secNewDate = moment(myNewDate).format("DD/MM");
          this.convert_axios2 = moment(myNewDate).format("YYYY-MM-DD");
         
          this.date_use_data.push({
            first_date:this.firstNewDate,
            sec_date:this.secNewDate,
            convert_axios1:this.convert_axios1,
            convert_axios2:this.convert_axios2
          })
          count++
       }else if(count == 1){
          
         this.next_date = myNewDate.setDate(myNewDate.getDate() - 12);
          this.firstNewDate = moment(myNewDate).format("DD/MM");
          this.convert_axios1 = moment(myNewDate).format("YYYY-MM-DD");
       

         
          this.next_date = myNewDate.setDate(myNewDate.getDate() + 5);
          this.secNewDate = moment(myNewDate).format("DD/MM");
          this.convert_axios2 = moment(myNewDate).format("YYYY-MM-DD");
         
          this.date_use_data.push({
            first_date:this.firstNewDate,
            sec_date:this.secNewDate,
            convert_axios1:this.convert_axios1,
            convert_axios2:this.convert_axios2
          })
          count++
       }else if(count == 2){
          
         this.next_date = myNewDate.setDate(myNewDate.getDate() - 12);
          this.firstNewDate = moment(myNewDate).format("DD/MM");
          this.convert_axios1 = moment(myNewDate).format("YYYY-MM-DD");
         

         
          this.next_date = myNewDate.setDate(myNewDate.getDate() + 5);
          this.secNewDate = moment(myNewDate).format("DD/MM");
          this.convert_axios2 = moment(myNewDate).format("YYYY-MM-DD");
       
          this.date_use_data.push({
            first_date:this.firstNewDate,
            sec_date:this.secNewDate,
            convert_axios1:this.convert_axios1,
            convert_axios2:this.convert_axios2
          })
          count++
       }else{
          
         this.next_date = myNewDate.setDate(myNewDate.getDate() - 12);
          this.firstNewDate = moment(myNewDate).format("DD/MM");
          this.convert_axios1 = moment(myNewDate).format("YYYY-MM-DD");
          

         
          this.next_date = myNewDate.setDate(myNewDate.getDate() + 5);
          this.secNewDate = moment(myNewDate).format("DD/MM");
          this.convert_axios2 = moment(myNewDate).format("YYYY-MM-DD");
        
          this.date_use_data.push({
            first_date:this.firstNewDate,
            sec_date:this.secNewDate,
            convert_axios1:this.convert_axios1,
            convert_axios2:this.convert_axios2
          })
          count++
       }
      }
     
      for(var ax = 3; ax>-1; ax--){
            this.date_use_data_re.push({
            first_date:this.date_use_data[ax].first_date,
            sec_date:this.date_use_data[ax].sec_date,
            convert_axios1:this.date_use_data[ax].convert_axios1,
            convert_axios2:this.date_use_data[ax].convert_axios2
            })
      }

      
      for(var ax = 0; ax<this.date_use_data.length; ax++){
        const params =  new FormData();
        params.append("start",this.date_use_data_re[ax].convert_axios1)
        params.append("end",this.date_use_data_re[ax].convert_axios2)
        let org = this.$q.localStorage.getItem("org");
        params.append("org",org)
       await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data_all_weekly_3_1',
        data:params,
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
              ttl_marker: e.TTL_MARKER_YD,
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
              balance: 
                e.ISSUE - (e.TTL_MARKER_YD * e.WEIGHT_YD) / 1000
              ,
            })
         
           
            
              
          });
          this.tll_marker_week = 0
          this.end_loss_week_1 = 0
          this.splice_week_1  = 0
          this.cut_week_1 = 0
          this.rem_week_1 = 0
          this.widthloss_week_1 = 0
        
        console.log(this.all_data_arr)
          
      for(var ax = 0; ax<this.all_data_arr.length; ax++){
        
      this.tll_marker_week = Number(this.tll_marker_week) + Number(this.all_data_arr[ax].ttl_marker)
      
      this.end_loss_week_1 = Number(this.end_loss_week_1) + Number(this.all_data_arr[ax].endless_length_yd)
      
      this.splice_week_1 = Number(this.splice_week_1) + Number(this.all_data_arr[ax].spliceloss)

      this.cut_week_1 = Number(this.cut_week_1) + Number(this.all_data_arr[ax].totalcutout)
      
      this.rem_week_1 = Number(this.rem_week_1) + Number(this.all_data_arr[ax].rement)
      

      
      
      }
      
      this.ttl_marker_yd.push({ //28
        value_ttl : this.tll_marker_week
      })

       this.end_week.push({       //29
        value_end : this.end_loss_week_1
      })
      
      
       this.splice_week.push({       //30
        value_splice : this.splice_week_1
      })

       this.cut_week.push({        //31
        value_cut : this.cut_week_1
      })

    
       this.rem_week.push({       //32
        value_rem : this.rem_week_1 
      })

     
        this.sum_widthloss = 0
        for(var ax = 0; ax<this.all_data_arr.length; ax++){
          this.widthloss_week_1 = 0
          this.widthloss_week_1 = this.all_data_arr[ax].ttl_marker * this.all_data_arr[ax].widthloss /100
          this.sum_widthloss = Number(this.sum_widthloss) + Number(this.widthloss_week_1)  
        }
        this.width_loss_yd.push({       //32
        value_width_loss : this.sum_widthloss
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
       this.ttl_marker_yd.push({
              value_ttl : 0
      })

       this.end_week.push({
        value_end : 0 
      })

       this.splice_week.push({
        value_splice : 0
      })

       this.cut_week.push({
        value_cut : 0
      })

       this.rem_week.push({
        value_rem : 0
      })
      this.width_loss_yd.push({       
        value_width_loss : 0
      })
     
       
      }
     
      
      })
      
      

      }
    
     
/// end 1-4 
/// start 52 week
var myNewDate = new Date(this.first_day);
var myNewDate2 = new Date(this.first_day);
         var count2 = 0;
for (var a1 = 1; a1 < 53; a1++) {
    
        
       if(count2 == 0 ){
          this.firstNewDate = moment(myNewDate2).format("DD/MM");
          this.convert_axios1 = moment(myNewDate2).format("YYYY-MM-DD");
         

          this.next_date =  myNewDate2.setDate(myNewDate2.getDate() + 5);
          this.secNewDate = moment(myNewDate2).format("DD/MM");
          this.convert_axios2 = moment(myNewDate2).format("YYYY-MM-DD");
       
          this.date_use_data2.push({
            first_date:this.firstNewDate,
            sec_date:this.secNewDate,
            convert_axios1:this.convert_axios1,
            convert_axios2:this.convert_axios2,
            count:a1
          })

          count2++
       }else{
          
         this.next_date = myNewDate2.setDate(myNewDate2.getDate() + 2);
          this.firstNewDate = moment(myNewDate2).format("DD/MM");
          this.convert_axios1 = moment(myNewDate2).format("YYYY-MM-DD");
       

         
          this.next_date = myNewDate2.setDate(myNewDate2.getDate() + 5);
          this.secNewDate = moment(myNewDate2).format("DD/MM");
          this.convert_axios2 = moment(myNewDate2).format("YYYY-MM-DD");
         
          this.date_use_data2.push({
            first_date:this.firstNewDate,
            sec_date:this.secNewDate,
            convert_axios1:this.convert_axios1,
            convert_axios2:this.convert_axios2,
            count:a1
          })
          count2++
       }
       }
       
       for(var ax = 0; ax<this.date_use_data2.length; ax++){
        const params2 = new FormData();
        params2.append("start",this.date_use_data2[ax].convert_axios1)
        params2.append("end",this.date_use_data2[ax].convert_axios2)
        let org = this.$q.localStorage.getItem("org");
        params2.append("org",org)
         await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data_all_type_3_1',
        data:params2,
      }).then((resp)=>{
    
         if (resp.data.data.length > 0) {
           this.all_data_arr2=[]
          resp.data.data.forEach((e) => {
            this.all_data_arr2.push(
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
              ttl_marker: e.TTL_MARKER_YD,
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
              balance: 
                e.ISSUE - (e.TTL_MARKER_YD * e.WEIGHT_YD) / 1000
              
            })
              
          });
          this.all_ttl_marker_yd = 0
          this.all_ac_end_loss = 0
          this.all_ac_splice = 0
          this.all_ac_cut = 0
          this.all_ac_rem = 0
          this.width_loss_yard = 0
          this.all_obswidth = 0
          this.all_sum_obs_width = 0
          this.total_obs_ttl = 0
          this.ttl_marker_end_woven = 0
          this.ttl_marker_end_light  = 0
          this.ttl_marker_end_fleece  = 0
          this.ttl_marker_end_special = 0
          this.ttl_marker_splice_woven2 = 0
          this.ttl_marker_splice_light2  = 0
          this.ttl_marker_splice_fleece2  = 0
          this.ttl_marker_splice_special2 = 0
          this.ttl_marker_cut_woven2 = 0
          this.ttl_marker_cut_light2 = 0
          this.ttl_marker_cut_fleece2  = 0
          this.ttl_marker_cut_special2 = 0
          this.ttl_marker_rem_woven2 = 0
          this.ttl_marker_rem_light2  = 0
          this.ttl_marker_rem_fleece2 = 0
          this.ttl_marker_rem_special2 = 0

          
          for(var ax = 0; ax<this.all_data_arr2.length; ax++){
              
              this.all_ttl_marker_yd = Number(this.all_ttl_marker_yd) + Number(this.all_data_arr2[ax].ttl_marker)

             this.all_ac_end_loss = Number(this.all_ac_end_loss) + Number(this.all_data_arr2[ax].endless_length_yd)
             
             this.all_ac_splice = Number(this.all_ac_splice) + Number(this.all_data_arr2[ax].spliceloss)

             this.all_ac_cut = Number(this.all_ac_cut) + Number(this.all_data_arr2[ax].totalcutout)

             this.all_ac_rem = Number(this.all_ac_rem) + Number(this.all_data_arr2[ax].rement)

             this.all_obswidth = Number(this.all_obswidth) + Number(this.all_data_arr2[ax].obswidth)

             this.total_obs_ttl = this.all_data_arr2[ax].obs_width * this.all_data_arr2[ax].ttl_marker

             this.all_sum_obs_width = Number(this.all_sum_obs_width) + Number(this.total_obs_ttl)
              }
            
            this.avg_width_loss_after_cal = 0
           
            this.avg_width_loss_after_cal = this.all_sum_obs_width /this.all_ttl_marker_yd

            this.avg_width_loss_cal.push({
              value:this.avg_width_loss_after_cal
            })


              this.ttl_marker_yd2.push({
                value_ttl : this.all_ttl_marker_yd
              })

              this.ac_end_loss2.push({
                  value_end : this.all_ac_end_loss
              })

              this.ac_splice2.push({
                  value_splice : this.all_ac_splice
              })

              this.ac_cut2.push({
                  value_cut : this.all_ac_cut
              })

              this.ac_rem2.push({
                  value_rem : this.all_ac_rem
              })
              this.ttl_obswidth.push({
                  value_obs:this.all_obswidth
              })
      //end
      for(var ax = 0; ax<this.all_data_arr2.length; ax++){
            if(this.all_data_arr2[ax].f_type == "1"){
              
            this.ttl_marker_end_woven = Number(this.ttl_marker_end_woven) + Number(this.all_data_arr2[ax].ttl_marker)
            
            }
           }
            this.ttl_marker_endless_woven.push({
              value : this.ttl_marker_end_woven,
            
            })
           
      
    for(var ax = 0; ax<this.all_data_arr2.length; ax++){
            if(this.all_data_arr2[ax].f_type == "2"){
              
            this.ttl_marker_end_light = Number(this.ttl_marker_end_light) + Number(this.all_data_arr2[ax].ttl_marker)
            
            }
          }
            this.ttl_marker_endless_light.push({
              value : this.ttl_marker_end_light,
              value_ax:ax
            })
          
          
for(var ax = 0; ax<this.all_data_arr2.length; ax++){
            if(this.all_data_arr2[ax].f_type == "3"){
              
            this.ttl_marker_end_fleece = Number(this.ttl_marker_end_fleece) + Number(this.all_data_arr2[ax].ttl_marker)
            
            
            }
            }
            this.ttl_marker_endless_fleece.push({
              value : this.ttl_marker_end_fleece
            })
            
          for(var ax = 0; ax<this.all_data_arr2.length; ax++){
            if(this.all_data_arr2[ax].f_type == "4"){
              
            this.ttl_marker_end_special = Number(this.ttl_marker_end_special) + Number(this.all_data_arr2[ax].ttl_marker)
            
            }
          }
          
            this.ttl_marker_endless_special.push({
              value : this.ttl_marker_end_special
            })
            

            //splice
          for(var ax = 0; ax<this.all_data_arr2.length; ax++){
            if(this.all_data_arr2[ax].f_type == "1"){
              
            this.ttl_marker_splice_woven2 = Number(this.ttl_marker_splice_woven2) + Number(this.all_data_arr2[ax].ttl_marker)
            
            }
            }
          
            this.ttl_marker_splice_woven.push({
              value : this.ttl_marker_splice_woven2
            })
            
            for(var ax = 0; ax<this.all_data_arr2.length; ax++){

            if(this.all_data_arr2[ax].f_type == "2"){
              
            this.ttl_marker_splice_light2 = Number(this.ttl_marker_splice_light2) + Number(this.all_data_arr2[ax].ttl_marker)
            }
            }
            
          
            this.ttl_marker_splice_light.push({
              value : this.ttl_marker_splice_light2
            })
            
        for(var ax = 0; ax<this.all_data_arr2.length; ax++){
            if(this.all_data_arr2[ax].f_type == "3"){
              
            this.ttl_marker_splice_fleece2 = Number(this.ttl_marker_splice_fleece2) + Number(this.all_data_arr2[ax].ttl_marker)
            
            }
            }
          
            this.ttl_marker_splice_fleece.push({
              value : this.ttl_marker_splice_fleece2
            })
            
          for(var ax = 0; ax<this.all_data_arr2.length; ax++){
            if(this.all_data_arr2[ax].f_type == "4"){
              
            this.ttl_marker_splice_special2 = Number(this.ttl_marker_splice_special2) + Number(this.all_data_arr2[ax].ttl_marker)
            
            }
            }
          
            this.ttl_marker_splice_special.push({
              value : this.ttl_marker_splice_special2
            })
            
 
 
            //cut
            for(var ax = 0; ax<this.all_data_arr2.length; ax++){
            if(this.all_data_arr2[ax].f_type == "1"){
              
            this.ttl_marker_cut_woven2 = Number(this.ttl_marker_cut_woven2) + Number(this.all_data_arr2[ax].ttl_marker)
            
            }
            }
          
            this.ttl_marker_cut_woven.push({
              value : this.ttl_marker_cut_woven2
            })
            
          for(var ax = 0; ax<this.all_data_arr2.length; ax++){
            if(this.all_data_arr2[ax].f_type == "2"){
              
            this.ttl_marker_cut_light2 = Number(this.ttl_marker_cut_light2) + Number(this.all_data_arr2[ax].ttl_marker)
            
            }
            }
          
            this.ttl_marker_cut_light.push({
              value : this.ttl_marker_cut_light2
            })
            
          for(var ax = 0; ax<this.all_data_arr2.length; ax++){
            if(this.all_data_arr2[ax].f_type == "3"){
              
            this.ttl_marker_cut_fleece2 = Number(this.ttl_marker_cut_fleece2) + Number(this.all_data_arr2[ax].ttl_marker)
            }
            }
            
          
            this.ttl_marker_cut_fleece.push({
              value : this.ttl_marker_cut_fleece2
            })
            
          for(var ax = 0; ax<this.all_data_arr2.length; ax++){
            if(this.all_data_arr2[ax].f_type == "4"){
              
            this.ttl_marker_cut_special2 = Number(this.ttl_marker_cut_special2) + Number(this.all_data_arr2[ax].ttl_marker)
            }
            }
            
          
            this.ttl_marker_cut_special.push({
              value : this.ttl_marker_cut_special2
            })
            

        //rem
        for(var ax = 0; ax<this.all_data_arr2.length; ax++){
         if(this.all_data_arr2[ax].f_type == "1"){
              
            this.ttl_marker_rem_woven2 = Number(this.ttl_marker_rem_woven2) + Number(this.all_data_arr2[ax].ttl_marker)
            
         }
         }
          
            this.ttl_marker_rem_woven.push({
              value : this.ttl_marker_rem_woven2
            })
            
        for(var ax = 0; ax<this.all_data_arr2.length; ax++){
            if(this.all_data_arr2[ax].f_type == "2"){
              
            this.ttl_marker_rem_light2 = Number(this.ttl_marker_rem_light2) + Number(this.all_data_arr2[ax].ttl_marker)
            }
        }
            
          
            this.ttl_marker_rem_light.push({
              value : this.ttl_marker_rem_light2
            })
            
          for(var ax = 0; ax<this.all_data_arr2.length; ax++){
            if(this.all_data_arr2[ax].f_type == "3"){
              
            this.ttl_marker_rem_fleece2 = Number(this.ttl_marker_rem_fleece2) + Number(this.all_data_arr2[ax].ttl_marker)
            }
          }
            
          
            this.ttl_marker_rem_fleece.push({
              value : this.ttl_marker_rem_fleece2
            })
            
        for(var ax = 0; ax<this.all_data_arr2.length; ax++){
            if(this.all_data_arr2[ax].f_type == "4"){
              
            this.ttl_marker_rem_special2 = Number(this.ttl_marker_rem_special2) + Number(this.all_data_arr2[ax].ttl_marker)
            }
            }
            
          
            this.ttl_marker_rem_special.push({
              value : this.ttl_marker_rem_special2
            })
            

            


              this.sum_widthloss2 = 0
        for(var ax = 0; ax<this.all_data_arr2.length; ax++){
          this.widthloss_week_2 = 0
          this.widthloss_week_2 = this.all_data_arr2[ax].ttl_marker * this.all_data_arr2[ax].widthloss /100
          this.sum_widthloss2 = Number(this.sum_widthloss2) + Number(this.widthloss_week_2)  
        }
        this.width_loss_yd_t.push({       //32
        value_width_loss : this.sum_widthloss2
      })       
             
      }else{
        this.all_data_arr2=[]
         this.all_data_arr2.push({
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
             this.ttl_marker_yd2.push({
                value_ttl : 0
            })
            this.ac_end_loss2.push({
                  value_end : 0
            })
            this.ac_splice2.push({
                  value_splice : 0
            })
            this.ac_cut2.push({
                  value_cut : 0
              })
            this.ac_rem2.push({
                  value_rem : 0
            })
              this.width_loss_yd_t.push({      
                 value_width_loss : 0
            })
            this.ttl_obswidth.push({
                  value_obs:0
              })

            this.ttl_marker_endless_woven.push({
              value : 0
            })

            this.ttl_marker_endless_light.push({
              value : 0
            })

            this.ttl_marker_endless_fleece.push({
              value : 0
            })

            this.ttl_marker_endless_special.push({
              value : 0
            })

             this.ttl_marker_splice_woven.push({
              value : 0
            })

            this.ttl_marker_splice_light.push({
              value : 0
            })

            this.ttl_marker_splice_fleece.push({
              value : 0
            })

            this.ttl_marker_splice_special.push({
              value : 0
            })

             this.ttl_marker_cut_woven.push({
              value : 0
            })

            this.ttl_marker_cut_light.push({
              value : 0
            })

            this.ttl_marker_cut_fleece.push({
              value : 0
            })

            this.ttl_marker_cut_special.push({
              value : 0
            })

             this.ttl_marker_rem_woven.push({
              value : 0
            })

            this.ttl_marker_rem_light.push({
              value : 0
            })

            this.ttl_marker_rem_fleece.push({
              value : 0
            })

            this.ttl_marker_rem_special.push({
              value : 0
            })
            this.avg_width_loss_cal.push({
              value: 0
            })

      }
        
      
      })
      
 
     

       }
        
       

        this.sum_all_ttl_marker = 0
        for(var ax =0; ax<this.ttl_marker_yd2.length; ax++){
           this.sum_all_ttl_marker = Number(this.sum_all_ttl_marker) + Number(this.ttl_marker_yd2[ax].value_ttl)
           
           this.ttl_marker_yd_t.push({
             value_ttl : this.sum_all_ttl_marker
           })
        }
       
       this.commulate_ttl = 0
       for(var ax = 0; ax<this.ttl_marker_yd2.length; ax++){
        this.commulate_ttl = Number(this.commulate_ttl) + Number(this.ttl_marker_yd2[ax].value_ttl)
        }
      
        this.sum_all_ac_end_loss = 0
        for(var ax =0; ax<this.ttl_marker_yd2.length; ax++){
           this.sum_all_ac_end_loss = Number(this.sum_all_ac_end_loss) + Number(this.ac_end_loss2[ax].value_end)
        }
       
        this.commulate_end = this.sum_all_ac_end_loss / this.commulate_ttl*100
        worksheet.getCell("N7").value = this.commulate_end/100
        worksheet.getCell("N7").numFmt = '0.00%'; 

        this.sum_all_ac_splice = 0
        for(var ax =0; ax<this.ttl_marker_yd2.length; ax++){
           this.sum_all_ac_splice = Number(this.sum_all_ac_splice) + Number(this.ac_splice2[ax].value_splice)
        }
      
        this.commulate_splice = this.sum_all_ac_splice / this.commulate_ttl*100
        worksheet.getCell("N8").value = this.commulate_splice/100
        worksheet.getCell("N8").numFmt = '0.00%'; 



        this.sum_all_ac_cut = 0
        for(var ax =0; ax<this.ttl_marker_yd2.length; ax++){
           this.sum_all_ac_cut = Number(this.sum_all_ac_cut) + Number(this.ac_cut2[ax].value_cut)
        }
        this.commulate_cut = this.sum_all_ac_cut / this.commulate_ttl*100
        worksheet.getCell("N9").value = this.commulate_cut/100
        worksheet.getCell("N9").numFmt = '0.00%'; 

      

     
        this.sum_all_ac_rem = 0
        for(var ax =0; ax<this.ttl_marker_yd2.length; ax++){
           this.sum_all_ac_rem = Number(this.sum_all_ac_rem) + Number(this.ac_rem2[ax].value_rem)
        }
        this.commulate_rem = this.sum_all_ac_rem / this.commulate_ttl*100
        worksheet.getCell("N10").value = this.commulate_rem/100
        worksheet.getCell("N10").numFmt = '0.00%'; 

    var date_1 = new Date(this.start_date)
    this.firstdate_1 = moment(date_1).format("DD/MM")
    var date_2 = new Date(this.end_date)
    this.lastdate_2 = moment(date_2).format("DD/MM")

   
      this.count_round_week = 0
     for(var ay = 0; ay<this.date_use_data2.length; ay++){
       console.log(this.date_use_data_re[3].first_date)
       console.log(this.date_use_data2[ay].first_date)
      if(this.date_use_data_re[3].first_date == this.date_use_data2[ay].first_date){
       this.count_round_week = this.date_use_data2[ay].count
     } 
    } 
    
    console.log(this.count_round_week)
    
    

        worksheet.getCell("K4").value = "this week"; //this.start - this.end


        for(var ax = 0; ax<52; ax++){
          if(this.firstdate_1 == this.date_use_data2[ax].first_date && this.lastdate_2 == this.date_use_data2[ax].sec_date){
            
            this.week = ax+1
            console.log(this.week)
            worksheet.getCell("J4").value = "Week:"+ax;
            console.log(ax-1)
            if(ax-1 <= 1){
              worksheet.getCell("I4").value = "week:";
              worksheet.getCell("H4").value = "week:"; 
              this.status_week_open = false
            }else{
            worksheet.getCell("I4").value = "week:"+(ax-1);
            worksheet.getCell("H4").value = "week:"+(ax-2); 
            }

            
            
          }
        }
       worksheet.getCell("A1").value = "Spread Loss Audit Weekly Summary ( MU-SLA-3-1 ) NY"+this.org+" Week  :"+this.week+" :  "+this.start+"-"+this.end; 
       worksheet.mergeCells("A1:O2");
 
    worksheet.getCell("O6").value = Number(this.commulate_ttl)
    worksheet.getCell("O6").numFmt = '#,##0.00';

    if(this.count_round_week == 1){
     if(this.status_week_open == false){
          for(var ax = 0; ax<this.ttl_marker_yd.length; ax++){
            if(ax == 0 || ax == 1 || ax == 2){
            this.ttl_marker_yd_change.push({
              value_ttl:""
            })
            }else{
              this.ttl_marker_yd_change.push({
              value_ttl:this.ttl_marker_yd[ax].value_ttl
            })
            }
          }
     }
     if(this.status_week_open == false){
          for(var ax = 0; ax<this.end_week.length; ax++){
            if(ax == 0 || ax == 1 || ax == 2){
            this.end_week_change.push({
              value_end:""
            })
            }else{
              this.end_week_change.push({
              value_end:this.end_week[ax].value_end
            })
            }
          }
     }

     if(this.status_week_open == false){
          for(var ax = 0; ax<this.splice_week.length; ax++){
            if(ax == 0 || ax == 1 || ax == 2){
            this.splice_week_change.push({
              value_splice:""
            })
            }else{
              this.splice_week_change.push({
              value_splice:this.splice_week[ax].value_splice
            })
            }
          }
     }

     if(this.status_week_open == false){
          for(var ax = 0; ax<this.cut_week.length; ax++){
            if(ax == 0 || ax == 1 || ax == 2){
            this.cut_week_change.push({
              value_cut:""
            })
            }else{
              this.cut_week_change.push({
              value_cut:this.cut_week[ax].value_cut
            })
            }
          }
     }

      if(this.status_week_open == false){
          for(var ax = 0; ax<this.rem_week.length; ax++){
            if(ax == 0 || ax == 1 || ax == 2){
            this.rem_week_change.push({
              value_rem:""
            })
            }else{
              this.rem_week_change.push({
              value_rem:this.rem_week[ax].value_rem
            })
            }
          }
     }
      if(this.status_week_open == false){
     for(var ax = 0; ax<this.width_loss_yd.length; ax++){
            if(ax == 0 || ax == 1 || ax == 2){
            this.width_loss_yd_change.push({
              value_width_loss:""
            })
            }else{
              this.width_loss_yd_change.push({
              value_width_loss:this.width_loss_yd[ax].value_width_loss
            })
            }
          }
      }

     if(this.status_week_open == false){
          for(var ax = 0; ax<this.width_loss_yd_t.length; ax++){
            if(ax == 0 || ax == 1 || ax == 2){
            this.width_loss_yd_t_change.push({
              value_rem:""
            })
            }else{
              this.width_loss_yd_t_change.push({
              value_rem:this.width_loss_yd_t[ax].value_width_loss 
            })
            }
          }
     }
    }else if(this.count_round_week == 2){
         if(this.status_week_open == false){
          for(var ax = 0; ax<this.ttl_marker_yd.length; ax++){
            if(ax == 0 || ax == 1 ){
            this.ttl_marker_yd_change.push({
              value_ttl:""
            })
            }else{
              this.ttl_marker_yd_change.push({
              value_ttl:this.ttl_marker_yd[ax].value_ttl
            })
            }
          }
     }
     if(this.status_week_open == false){
          for(var ax = 0; ax<this.end_week.length; ax++){
            if(ax == 0 || ax == 1 ){
            this.end_week_change.push({
              value_end:""
            })
            }else{
              this.end_week_change.push({
              value_end:this.end_week[ax].value_end
            })
            }
          }
     }

     if(this.status_week_open == false){
          for(var ax = 0; ax<this.splice_week.length; ax++){
            if(ax == 0 || ax == 1 ){
            this.splice_week_change.push({
              value_splice:""
            })
            }else{
              this.splice_week_change.push({
              value_splice:this.splice_week[ax].value_splice
            })
            }
          }
     }

     if(this.status_week_open == false){
          for(var ax = 0; ax<this.cut_week.length; ax++){
            if(ax == 0 || ax == 1 ){
            this.cut_week_change.push({
              value_cut:""
            })
            }else{
              this.cut_week_change.push({
              value_cut:this.cut_week[ax].value_cut
            })
            }
          }
     }

      if(this.status_week_open == false){
          for(var ax = 0; ax<this.rem_week.length; ax++){
            if(ax == 0 || ax == 1 ){
            this.rem_week_change.push({
              value_rem:""
            })
            }else{
              this.rem_week_change.push({
              value_rem:this.rem_week[ax].value_rem
            })
            }
          }
     }

      if(this.status_week_open == false){
     for(var ax = 0; ax<this.width_loss_yd.length; ax++){
            if(ax == 0 || ax == 1){
            this.width_loss_yd_change.push({
              value_width_loss:""
            })
            }else{
              this.width_loss_yd_change.push({
              value_width_loss:this.width_loss_yd[ax].value_width_loss
            })
            }
          }
      }
      if(this.status_week_open == false){
          for(var ax = 0; ax<this.width_loss_yd_t.length; ax++){
            if(ax == 0 || ax == 1 ){
            this.width_loss_yd_t_change.push({
              value_rem:""
            })
            }else{
              this.width_loss_yd_t_change.push({
              value_rem:this.width_loss_yd_t[ax].value_width_loss 
            })
            }
          }
     }
      }else if(this.count_round_week == 3){
          if(this.status_week_open == false){
          for(var ax = 0; ax<this.ttl_marker_yd.length; ax++){
            if(ax == 0){
            this.ttl_marker_yd_change.push({
              value_ttl:""
            })
            }else{
              this.ttl_marker_yd_change.push({
              value_ttl:this.ttl_marker_yd[ax].value_ttl
            })
            }
          }
     }
     if(this.status_week_open == false){
          for(var ax = 0; ax<this.end_week.length; ax++){
            if(ax == 0){
            this.end_week_change.push({
              value_end:""
            })
            }else{
              this.end_week_change.push({
              value_end:this.end_week[ax].value_end
            })
            }
          }
     }

     if(this.status_week_open == false){
          for(var ax = 0; ax<this.splice_week.length; ax++){
            if(ax == 0){
            this.splice_week_change.push({
              value_splice:""
            })
            }else{
              this.splice_week_change.push({
              value_splice:this.splice_week[ax].value_splice
            })
            }
          }
     }

     if(this.status_week_open == false){
          for(var ax = 0; ax<this.cut_week.length; ax++){
            if(ax == 0){
            this.cut_week_change.push({
              value_cut:""
            })
            }else{
              this.cut_week_change.push({
              value_cut:this.cut_week[ax].value_cut
            })
            }
          }
     }

      if(this.status_week_open == false){
          for(var ax = 0; ax<this.rem_week.length; ax++){
            if(ax == 0){
            this.rem_week_change.push({
              value_rem:""
            })
            }else{
              this.rem_week_change.push({
              value_rem:this.rem_week[ax].value_rem
            })
            }
          }
     }

      if(this.status_week_open == false){
     for(var ax = 0; ax<this.width_loss_yd.length; ax++){
            if(ax == 0){
            this.width_loss_yd_change.push({
              value_width_loss:""
            })
            }else{
              this.width_loss_yd_change.push({
              value_width_loss:this.width_loss_yd[ax].value_width_loss
            })
            }
          }
      }

      if(this.status_week_open == false){
          for(var ax = 0; ax<this.width_loss_yd_t.length; ax++){
            if(ax == 0 || ax == 1 || ax == 2){
            this.width_loss_yd_t_change.push({
              value_rem:""
            })
            }else{
              this.width_loss_yd_t_change.push({
              value_rem:this.width_loss_yd_t[ax].value_width_loss 
            })
            }
          }
     }

    }

 
      
   
    for(var ax = 0; ax<this.ttl_marker_yd.length; ax++){
    
     worksheet.getCell(this.colume_arr[ax].cn+"5").value = this.date_use_data_re[ax].first_date +" - "+ this.date_use_data_re[ax].sec_date;
   
    if(this.status_week_open == false){
      worksheet.getCell(this.colume_arr[ax].cn+"6").value = Number(this.ttl_marker_yd_change[ax].value_ttl);
    worksheet.getCell(this.colume_arr[ax].cn+"6").numFmt = '#,##0.00';
    }else{
    worksheet.getCell(this.colume_arr[ax].cn+"6").value = Number(this.ttl_marker_yd[ax].value_ttl);
    worksheet.getCell(this.colume_arr[ax].cn+"6").numFmt = '#,##0.00';
    }
    

    }

    
    

    this.comulate_total_ac = 0
       
       for(var ax = 0; ax<this.ttl_marker_yd.length; ax++){
         this.endloss_ac = 0
         if(this.status_week_open == false){
           this.endloss_ac = this.end_week_change[ax].value_end/this.ttl_marker_yd_change[ax].value_ttl*100
         }else{
        this.endloss_ac = this.end_week[ax].value_end/this.ttl_marker_yd[ax].value_ttl*100
         }
         
        
         if(isNaN(this.endloss_ac) == true){
         this.endloss_ac = 0
         }
        
         
          
         this.val_end.push({
           value_end: this.endloss_ac
         })
         if(this.endloss_ac > 0){
         worksheet.getCell(this.colume_arr[ax].cn+"7").value = this.endloss_ac/100;
         worksheet.getCell(this.colume_arr[ax].cn+"7").numFmt = '0.00%'; 
         }else{
          worksheet.getCell(this.colume_arr[ax].cn+"7").value = 0.00;
          worksheet.getCell(this.colume_arr[ax].cn+"7").numFmt = "0.00%";
         } 
       


       
         this.splice_ac = 0
         if(this.status_week_open == false){
           this.splice_ac = this.splice_week_change[ax].value_splice/this.ttl_marker_yd_change[ax].value_ttl*100
         }else{
           this.splice_ac = this.splice_week[ax].value_splice/this.ttl_marker_yd[ax].value_ttl*100
         }
      
         
         if(isNaN(this.splice_ac) == true){
         this.splice_ac = 0
         }
        
        
         this.val_splice.push({
           value_splice: this.splice_ac
         })
         if(this.splice_ac > 0){
         worksheet.getCell(this.colume_arr[ax].cn+"8").value = this.splice_ac/100;
         worksheet.getCell(this.colume_arr[ax].cn+"8").numFmt = "0.00%";
         
         }else{
          worksheet.getCell(this.colume_arr[ax].cn+"8").value = 0.00; 
          worksheet.getCell(this.colume_arr[ax].cn+"8").numFmt = "0.00%";
         } 

         

         
      
  
       
         this.cut_ac = 0
         if(this.status_week_open == false){
this.cut_ac = this.cut_week_change[ax].value_cut/this.ttl_marker_yd_change[ax].value_ttl*100
         }else{
this.cut_ac = this.cut_week[ax].value_cut/this.ttl_marker_yd[ax].value_ttl*100
         }
         
         if(isNaN(this.cut_ac) == true){
         this.cut_ac = 0
         }
        
         this.val_cut.push({
           value_cut: this.cut_ac
         })
         if(this.cut_ac > 0){
         worksheet.getCell(this.colume_arr[ax].cn+"9").value = this.cut_ac/100
         worksheet.getCell(this.colume_arr[ax].cn+"9").numFmt = "0.00%";
         }else{
          worksheet.getCell(this.colume_arr[ax].cn+"9").value = 0.00;
          worksheet.getCell(this.colume_arr[ax].cn+"9").numFmt = "0.00%";
         } 
       
     
         this.rem_ac = 0
         if(this.status_week_open == false){
            this.rem_ac = this.rem_week_change[ax].value_rem/this.ttl_marker_yd_change[ax].value_ttl*100
         }else{
            this.rem_ac = this.rem_week[ax].value_rem/this.ttl_marker_yd[ax].value_ttl*100
         }
         
         if(isNaN(this.rem_ac) == true){
         this.rem_ac = 0
         }
        
        
          this.val_rem.push({
           value_rem: this.rem_ac
         })
         if(this.rem_ac > 0){
         worksheet.getCell(this.colume_arr[ax].cn+"10").value = this.rem_ac/100
         worksheet.getCell(this.colume_arr[ax].cn+"10").numFmt = "0.00%";
         }else{
          worksheet.getCell(this.colume_arr[ax].cn+"10").value = 0.00
          worksheet.getCell(this.colume_arr[ax].cn+"10").numFmt = "0.00%";
          
        }
     

         this.total_ac = 0
         this.total_ac = Number(this.val_end[ax].value_end) + Number(this.val_splice[ax].value_splice)+
          Number(this.val_cut[ax].value_cut) + Number(this.val_rem[ax].value_rem) 
       
          if(this.total_ac > 0){
         worksheet.getCell(this.colume_arr[ax].cn+"11").value = this.total_ac/100;
         worksheet.getCell(this.colume_arr[ax].cn+"11").numFmt = "0.00%";
         }else{
          worksheet.getCell(this.colume_arr[ax].cn+"11").value = 0.00; 
          worksheet.getCell(this.colume_arr[ax].cn+"11").numFmt = "0.00%";
         } 
         
         this.comulate_total_ac = Number(this.comulate_total_ac)+Number(this.total_ac)

         
       }

       
         this.commulate_total_all_loss = 0
       for(var ax = 0; ax<this.colume_arr.length; ax++){  //33
         this.total_all_loss = 0
         if(this.status_week_open == false){
           this.total_all_loss = Number(this.end_week_change[ax].value_end) + Number(this.splice_week_change[ax].value_splice) 
         +Number(this.cut_week_change[ax].value_cut) + Number(this.rem_week_change[ax].value_rem)
         }else{
         this.total_all_loss = Number(this.end_week[ax].value_end) + Number(this.splice_week[ax].value_splice) 
         +Number(this.cut_week[ax].value_cut) + Number(this.rem_week[ax].value_rem)
         }

           worksheet.getCell(this.colume_arr[ax].cn+"12").value = Number(this.total_all_loss)
           worksheet.getCell(this.colume_arr[ax].cn+"12").numFmt = "#,##0.00";
           
            this.commulate_total_all_loss = Number(this.commulate_total_all_loss) + Number(this.total_all_loss)

       }
            

             worksheet.getCell("M12").value = Number(this.commulate_total_all_loss )
             worksheet.getCell("M12").numFmt = '0.00'; 
      
    
        this.sum_all_ttl_marker_4_week = 0

        for(var ax = 0; ax<this.ttl_marker_yd.length; ax++){
          if(this.status_week_open == false){
          this.sum_all_ttl_marker_4_week = Number(this.sum_all_ttl_marker_4_week) + Number(this.ttl_marker_yd_change[ax].value_ttl)
          }else{
          this.sum_all_ttl_marker_4_week = Number(this.sum_all_ttl_marker_4_week) + Number(this.ttl_marker_yd[ax].value_ttl)
          }
        }
          worksheet.getCell("M6").value = Number(this.sum_all_ttl_marker_4_week);
          worksheet.getCell("M6").numFmt = '#,##0.00'; 




         worksheet.getCell("L11").value = this.commulate_total_all_loss/this.sum_all_ttl_marker_4_week 
         worksheet.getCell("L11").numFmt = '0.00%'; 
        
        this.end_loss_4_week = 0
        this.splice_loss_4_week = 0
        this.cut_loss_4_week = 0
        this.rem_loss_4_week = 0
        
        for(var ax = 0; ax<this.end_week.length; ax++){
          if(this.status_week_open == false){
          this.end_loss_4_week = Number(this.end_loss_4_week) + Number(this.end_week_change[ax].value_end) 
          this.splice_loss_4_week = Number(this.splice_loss_4_week) + Number(this.splice_week_change[ax].value_splice)
          this.cut_loss_4_week = Number(this.cut_loss_4_week) + Number(this.cut_week_change[ax].value_cut)
          this.rem_loss_4_week = Number(this.rem_loss_4_week) + Number(this.rem_week_change[ax].value_rem)
          }else{
          this.end_loss_4_week = Number(this.end_loss_4_week) + Number(this.end_week[ax].value_end) 
          this.splice_loss_4_week = Number(this.splice_loss_4_week) + Number(this.splice_week[ax].value_splice)
          this.cut_loss_4_week = Number(this.cut_loss_4_week) + Number(this.cut_week[ax].value_cut)
          this.rem_loss_4_week = Number(this.rem_loss_4_week) + Number(this.rem_week[ax].value_rem)
          }
        }
     

        
          this.end_loss_4_week = this.end_loss_4_week / this.sum_all_ttl_marker_4_week*100
          worksheet.getCell("L7").value = this.end_loss_4_week/100;
          worksheet.getCell("L7").numFmt = '0.00%'; 
        

          this.splice_loss_4_week = this.splice_loss_4_week / this.sum_all_ttl_marker_4_week*100
          worksheet.getCell("L8").value = this.splice_loss_4_week/100
          worksheet.getCell("L8").numFmt = '0.00%';
        

          this.cut_loss_4_week  = this.cut_loss_4_week / this.sum_all_ttl_marker_4_week*100
          worksheet.getCell("L9").value = this.cut_loss_4_week/100
          worksheet.getCell("L9").numFmt = '0.00%';
     

          this.rem_loss_4_week = this.rem_loss_4_week / this.sum_all_ttl_marker_4_week *100
          worksheet.getCell("L10").value = this.rem_loss_4_week/100
          worksheet.getCell("L10").numFmt = '0.00%'; 
     


          worksheet.getCell("S1").value = ""; 
       worksheet.mergeCells("S1:S80");
        worksheet.getCell("S1").fill = {
          type: 'pattern',
          pattern:'solid',
          fgColor:{argb:'FF880000'},
          bgColor:{argb:'FF880000'}
          };

        for(var ax = 0; ax<this.arr_colume2.length; ax++){
              worksheet.getCell(this.arr_colume2[ax].name_col+"4").value = "Week"+[ax+1];
              worksheet.getCell(this.arr_colume2[ax].name_col+"5").value = this.date_use_data2[ax].first_date + "-" +this.date_use_data2[ax].sec_date;
            
        }
        

     

         worksheet.getCell("R6").value = "Total  Marker  (Yds)";

         for(var ax = 0; ax<this.arr_colume2.length; ax++){

              if(this.ttl_marker_yd2[ax].value_ttl > 0){
              worksheet.getCell(this.arr_colume2[ax].name_col+"6").value = Number(this.ttl_marker_yd2[ax].value_ttl);
              worksheet.getCell(this.arr_colume2[ax].name_col+"6").numFmt = '#,##0.00';
              
              }else{
                worksheet.getCell(this.arr_colume2[ax].name_col+"6").value = Number(0);
                worksheet.getCell(this.arr_colume2[ax].name_col+"6").numFmt = '0.00';
                
              }
              
        }
         worksheet.getCell("R7").value = "End Loss";

         for(var ax = 0; ax<this.arr_colume2.length; ax++){
          // this.end_loss_light_per = this.ttl_endloss_yd_light[ax].end_loss_yd_light/this.ttl_marker_endless_light[ax].end_ttl_marker_light*100
          this.end_per = this.ac_end_loss2[ax].value_end / this.ttl_marker_yd2[ax].value_ttl *100
            if(this.end_per > 0){
              worksheet.getCell(this.arr_colume2[ax].name_col+"7").value = this.end_per/100;
              worksheet.getCell(this.arr_colume2[ax].name_col+"7").numFmt = "0.00%";
            }else{
              worksheet.getCell(this.arr_colume2[ax].name_col+"7").value = 0.00;
              worksheet.getCell(this.arr_colume2[ax].name_col+"7").numFmt = "0.00%";
            }
              
        }
         worksheet.getCell("R8").value = "Splice Loss";
         for(var ax = 0; ax<this.arr_colume2.length; ax++){
          this.splice_per = this.ac_splice2[ax].value_splice/ this.ttl_marker_yd2[ax].value_ttl *100
          if(this.splice_per > 0){
              worksheet.getCell(this.arr_colume2[ax].name_col+"8").value = this.splice_per/100;
              worksheet.getCell(this.arr_colume2[ax].name_col+"8").numFmt = "0.00%";
          }else{
              worksheet.getCell(this.arr_colume2[ax].name_col+"8").value = 0.00;
              worksheet.getCell(this.arr_colume2[ax].name_col+"8").numFmt = "0.00%";
          }       
        }

        
         worksheet.getCell("R9").value = "Cut out Loss";

           for(var ax = 0; ax<this.arr_colume2.length; ax++){
          this.cut_per = this.ac_cut2[ax].value_cut/ this.ttl_marker_yd2[ax].value_ttl *100
          if(this.cut_per > 0){
              worksheet.getCell(this.arr_colume2[ax].name_col+"9").value = this.cut_per/100;
              worksheet.getCell(this.arr_colume2[ax].name_col+"9").numFmt = "0.00%";
          }else{
              worksheet.getCell(this.arr_colume2[ax].name_col+"9").value = 0.00;
              worksheet.getCell(this.arr_colume2[ax].name_col+"9").numFmt = "0.00%";
          }       
        }

         worksheet.getCell("R10").value = "Remnants Loss";

            for(var ax = 0; ax<this.arr_colume2.length; ax++){
          this.rem_per = this.ac_rem2[ax].value_rem/ this.ttl_marker_yd2[ax].value_ttl *100
          if(this.rem_per > 0){
              worksheet.getCell(this.arr_colume2[ax].name_col+"10").value = this.rem_per/100
              worksheet.getCell(this.arr_colume2[ax].name_col+"10").numFmt = "0.00%";
          }else{
              worksheet.getCell(this.arr_colume2[ax].name_col+"10").value = 0.00;
              worksheet.getCell(this.arr_colume2[ax].name_col+"10").numFmt = "0.00%";
          }       
        }
         worksheet.getCell("R11").value = "% Total  Spread  Loss";
         worksheet.getCell("R12").value = "Total  Spread  Loss  (Yds)";
         worksheet.getCell("R13").value = "Total Width Loss(Yds) ";
         worksheet.getCell("R14").value = "Total Width Loss (%)";
          worksheet.getCell("R15").value = "Total Fabric Cut (Yds)";
  
      
          this.commulate_yds_width_loss = 0
          this.cummulate_4_week_ttl_yd = 0
          for(var ax = 0; ax<this.arr_colume2.length; ax++){

             
              this.total_spread_loss_yd = 0
              this.total_spread_loss_yd = Number(this.ac_end_loss2[ax].value_end) + Number(this.ac_splice2[ax].value_splice)
              +Number(this.ac_cut2[ax].value_cut) + Number(this.ac_rem2[ax].value_rem)
              this.commulate_spread_yd.push({
                value_yd:this.total_spread_loss_yd
              })
                if(this.total_spread_loss_yd > 0){
                  worksheet.getCell(this.arr_colume2[ax].name_col+"12").value = Number(this.total_spread_loss_yd);
                  worksheet.getCell(this.arr_colume2[ax].name_col+"12").numFmt = "#,##0.00";
                }else{
                  worksheet.getCell(this.arr_colume2[ax].name_col+"12").value = 0.00;
                  worksheet.getCell(this.arr_colume2[ax].name_col+"12").numFmt = "0.00";
                }

                this.commulate_yds_width_loss = Number(this.commulate_yds_width_loss) + Number(this.total_spread_loss_yd)
              
              this.total_spread_loss = 0
    
              this.total_spread_loss = this.total_spread_loss_yd / this.ttl_marker_yd2[ax].value_ttl
             
              if(isNaN(this.total_spread_loss) == true){
                this.total_spread_loss = 0
              }
              
              this.total_spread_width_loss.push({
                value_width_loss : this.width_loss_yd_t[ax].value_width_loss
              })   

              if(this.total_spread_loss > 0){
                  worksheet.getCell(this.arr_colume2[ax].name_col+"13").value = Number(this.width_loss_yd_t[ax].value_width_loss);
                  worksheet.getCell(this.arr_colume2[ax].name_col+"13").numFmt = "#,##0.00";
                }else{
                  worksheet.getCell(this.arr_colume2[ax].name_col+"13").value = 0;
                  worksheet.getCell(this.arr_colume2[ax].name_col+"13").numFmt = "0.00";
                }


   

                this.total_spread_loss_t = 0
                this.total_spread_loss_t =   this.commulate_spread_yd[ax].value_yd / this.ttl_marker_yd2[ax].value_ttl *100
               
                if(isNaN(this.total_spread_loss_t )== true){
                  this.total_spread_loss_t = 0
                }
            this.ttl_spread_graph.push({
             value:this.total_spread_loss_t 
           })
                
                if(this.total_spread_loss_t > 0){
                  worksheet.getCell(this.arr_colume2[ax].name_col+"11").value = this.total_spread_loss_t/100
                  worksheet.getCell(this.arr_colume2[ax].name_col+"11").numFmt = "0.00%";
                }else{
                  worksheet.getCell(this.arr_colume2[ax].name_col+"11").value = 0.00;
                  worksheet.getCell(this.arr_colume2[ax].name_col+"11").numFmt = "0.00%";
                }
              }

          
            this.cummulate_4_week_ttl_yd = 0
            for (var ax = 0; ax<this.width_loss_yd.length; ax++){
              if(this.status_week_open == false){
                this.cummulate_4_week_ttl_yd = Number(this.cummulate_4_week_ttl_yd) + Number(this.width_loss_yd_change[ax].value_width_loss) 
              }else{
                this.cummulate_4_week_ttl_yd = Number(this.cummulate_4_week_ttl_yd) + Number(this.width_loss_yd[ax].value_width_loss)
              }
            }
           
              worksheet.getCell("M13").value = Number(this.cummulate_4_week_ttl_yd)
              worksheet.getCell("M13").numFmt = '0.00'; 
             
             

             this.total_width_loss_4_week = 0
           
             this.total_width_loss_4_week = this.cummulate_4_week_ttl_yd / this.sum_all_ttl_marker_4_week *100

             worksheet.getCell("L14").value = this.total_width_loss_4_week/100
              worksheet.getCell("L14").numFmt = '0.00%';

          


          //this.total_spread_loss_yd
        

          for(var ax = 0; ax<this.arr_colume2.length; ax++){
           this.ac_width = 0
          
           this.ac_width = this.total_spread_width_loss[ax].value_width_loss / this.ttl_marker_yd2[ax].value_ttl *100
          
            if(this.ac_width  > 0){
                  worksheet.getCell(this.arr_colume2[ax].name_col+"14").value = this.ac_width/100;
                  worksheet.getCell(this.arr_colume2[ax].name_col+"14").numFmt = "0.00%";
                }else{
                  worksheet.getCell(this.arr_colume2[ax].name_col+"14").value = 0.00;
                  worksheet.getCell(this.arr_colume2[ax].name_col+"14").numFmt = "0.00%";
                }      
         }
         
         //cut 4 week
         
         for(var ax = 0; ax<this.date_use_data.length; ax++){  // Fabric cut yard 4 week
     const params6 = new FormData();
     
        params6.append("start",this.date_use_data_re[ax].convert_axios1)
        params6.append("end",this.date_use_data_re[ax].convert_axios2)
        let org = this.$q.localStorage.getItem("org");
        params6.append("org",org)
  await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data_all_weekly_yard',
        data:params6,
      }).then((resp)=>{
        
        
        if(resp.data.data[0].FABRIC_WEIGHT !== null){
          resp.data.data.forEach((e) => {
            this.all_week_cut_4.push(
              {
              value_cut_yard:e.FABRIC_WEIGHT
            })

          });
        }else{
            this.all_week_cut_4.push({
              value_cut_yard:"0"
            })
        }
        
      }) 
      
  }


  for(var ax = 0; ax<this.date_use_data.length; ax++){ //Dozen Cut แบบ 4 week
     const params7 = new FormData();
        params7.append("start",this.date_use_data_re[ax].convert_axios1)
        params7.append("end",this.date_use_data_re[ax].convert_axios2)
        let org = this.$q.localStorage.getItem("org");
        params7.append("org",org)
  await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data_all_weekly_cut',
        data:params7,
      }).then((resp)=>{
        
        
         if(resp.data.data[0].CUT_QTY !== null){
          resp.data.data.forEach((e) => {
            this.all_week_cut4_4.push(
              {
              value_cut:e.CUT_QTY
            })

          });
        }else if(resp.data.data[0].CUT_QTY == null){
            this.all_week_cut4_4.push({
              value_cut:"0"
            })
        }
        
      })
  }
             
             //Cut All Day
  for(var ax = 0; ax<this.date_use_data2.length; ax++){   //Total Dozens Cut / 12 all week
     const params3 = new FormData();
        params3.append("start",this.date_use_data2[ax].convert_axios1)
        params3.append("end",this.date_use_data2[ax].convert_axios2)
        let org = this.$q.localStorage.getItem("org");
        params3.append("org",org)
  await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data_all_weekly_cut',
        data:params3,
      }).then((resp)=>{
        
       
         if(resp.data.data[0].CUT_QTY !== null){
          resp.data.data.forEach((e) => {
            this.all_week_yard.push(
              {
              value_cut_yard:e.CUT_QTY
            })

          });
        }else if(resp.data.data[0].CUT_QTY == null){
            this.all_week_yard.push({
              value_cut_yard:"0"
            })
        }
      
      })
  }
 for(var ax = 0; ax<this.date_use_data2.length; ax++){ //Total Fabric Cut (Yds) all week (2)
    const params4 = new FormData();
        params4.append("start",this.date_use_data2[ax].convert_axios1)
        params4.append("end",this.date_use_data2[ax].convert_axios2)
        let org = this.$q.localStorage.getItem("org");
        params4.append("org",org)
       await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data_all_weekly_yard',
        data:params4,
      }).then((resp)=>{
       
       
        if(resp.data.data[0].FABRIC_WEIGHT !== null){
          resp.data.data.forEach((e) => {
            this.all_week_cut.push(
              {
              value_cut_yard:e.FABRIC_WEIGHT
            })

          });
        }else{
            this.all_week_cut.push({
              value_cut_yard:"0"
            })
        }
       
      }) 
 }

  for(var ax = 0; ax<this.date_use_data2.length; ax++){ //Total Fabric Cut (Yds) all week (3)
    const params14 = new FormData();
        params14.append("start",this.date_use_data2[ax].convert_axios1)
        params14.append("end",this.date_use_data2[ax].convert_axios2)
        let org = this.$q.localStorage.getItem("org");
        params14.append("org",org)
       await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data_all_weekly_yard_2',
        data:params14,
      }).then((resp)=>{
        
       
        if(resp.data.data[0].FABRIC_WEIGHT !== null){
          resp.data.data.forEach((e) => {
            this.all_week_cut_2.push(
              {
              value_cut_yard:e.FABRIC_WEIGHT
            })

          });
        }else{
            this.all_week_cut_2.push({
              value_cut_yard:"0"
            })
        }
       
      }) 
 }


 for(var ax = 0; ax<this.date_use_data.length; ax++){ // Total Fabric Cut (Yds) 4 week(2)
     const params7 = new FormData();
        params7.append("start",this.date_use_data_re[ax].convert_axios1)
        params7.append("end",this.date_use_data_re[ax].convert_axios2)
        let org = this.$q.localStorage.getItem("org");
        params7.append("org",org)
  await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data_all_weekly_yard',
        data:params7,
      }).then((resp)=>{
        
        
         if(resp.data.data[0].FABRIC_WEIGHT !== null){
          resp.data.data.forEach((e) => {
            this.all_week_cut_yard_4_4_1.push(
              { 
              value_cut:e.FABRIC_WEIGHT
            })

          });
        }else if(resp.data.data[0].FABRIC_WEIGHT == null){
            this.all_week_cut_yard_4_4_1.push({
              value_cut:"0"
            })
        }
        
      })
  }

  for(var ax = 0; ax<this.date_use_data.length; ax++){ // Total Fabric Cut (Yds) 4 week(3)
     const params9 = new FormData();
        params9.append("start",this.date_use_data_re[ax].convert_axios1)
        params9.append("end",this.date_use_data_re[ax].convert_axios2)
        let org = this.$q.localStorage.getItem("org");
        params9.append("org",org)
  await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data_all_weekly_yard_2',
        data:params9,
      }).then((resp)=>{
       
        
         if(resp.data.data[0].FABRIC_WEIGHT !== null){
          resp.data.data.forEach((e) => {
            this.all_week_cut_yard_4_4_2.push(
              { 
              value_cut:e.FABRIC_WEIGHT
            })

          });
        }else if(resp.data.data[0].FABRIC_WEIGHT == null){
            this.all_week_cut_yard_4_4_2.push({
              value_cut:"0"
            })
        }
        
      })
  }
      


        this.commulate_total_fabric_cut = 0
        this.sum_week_yard = 0
         for(var ax = 0; ax<this.arr_colume2.length; ax++){
           
           this.sum_week_yard = Number(this.all_week_cut[ax].value_cut_yard)  + Number(this.all_week_cut_2[ax].value_cut_yard)
            if(this.sum_week_yard  > 0){
                  worksheet.getCell(this.arr_colume2[ax].name_col+"15").value = Number(this.sum_week_yard);
                  worksheet.getCell(this.arr_colume2[ax].name_col+"15").numFmt = "#,##0.00";
                }else{
                  worksheet.getCell(this.arr_colume2[ax].name_col+"15").value = 0.00;
                  worksheet.getCell(this.arr_colume2[ax].name_col+"15").numFmt = "0.00";
                }
              
               this.commulate_total_fabric_cut = Number(this.commulate_total_fabric_cut) + Number(this.sum_week_yard)
            
         }    

          worksheet.getCell("O15").value = Number(this.commulate_total_fabric_cut);
          worksheet.getCell("O15").numFmt = "#,##0.00";
          this.sum_commulate_spread_yd = 0
          for(var ax = 0; ax<this.commulate_spread_yd.length; ax++){
            this.sum_commulate_spread_yd = Number(this.sum_commulate_spread_yd) + Number(this.commulate_spread_yd[ax].value_yd)
          }
          worksheet.getCell("O12").value = Number(this.sum_commulate_spread_yd);
          worksheet.getCell("O12").numFmt = "#,##0.00";


              this.cumlate_4_total = 0
   
             this.cumlate_4_total = this.sum_commulate_spread_yd / this.commulate_ttl
            
              worksheet.getCell("N11").value = this.cumlate_4_total;
               worksheet.getCell("N11").numFmt = "0.00%";


    

        this.sum_commulate_widthloss = 0
          for(var ax = 0; ax<this.width_loss_yd_t.length; ax++){
            this.sum_commulate_widthloss = Number(this.sum_commulate_widthloss) + Number(this.width_loss_yd_t[ax].value_width_loss)
          }
          worksheet.getCell("O13").value = Number(this.sum_commulate_widthloss);
           worksheet.getCell("O13").numFmt = "#,##0.00";


    
        this.total_spread_loss_per = this.sum_commulate_widthloss/this.commulate_ttl*100

        worksheet.getCell("N14").value = this.total_spread_loss_per/100;
        worksheet.getCell("N14").numFmt = "0.00%";

      this.sum_all_week_yard = 0
       for(var ax = 0; ax<this.width_loss_yd_t.length; ax++){
       this.sum_all_week_yard = Number(this.sum_all_week_yard) + Number(this.all_week_yard[ax].value_cut_yard)
       }
        
      this.commulate_total_fabric_cut
      this.sum_all_week_yard_per = 0
      this.sum_all_week_yard_per = this.commulate_ttl / this.commulate_total_fabric_cut

      worksheet.getCell("N16").value = this.sum_all_week_yard_per;
      worksheet.getCell("N16").numFmt = "0.00%";
      //Weight Target of the week
       //endloss
       
          this.actual_per_woven_end = this.end_loss_woven * this.ttl_marker_1_4[0].value
     
          this.actual_per_light_end = this.end_loss_light * this.ttl_marker_1_4[1].value

          this.actual_per_fleece_end = this.end_loss_fleece * this.ttl_marker_1_4[2].value
    
          this.actual_per_special_end = this.end_loss_special * this.ttl_marker_1_4[3].value
        
          
          this.all_actual_endloss = Number(this.actual_per_woven_end) + Number(this.actual_per_light_end) + Number(this.actual_per_fleece_end) + Number(this.actual_per_special_end)
          
          this.all_actual_endloss = this.all_actual_endloss / this.ttl_marker_1_4_all[0].value
   

            worksheet.getCell("G7").value = this.all_actual_endloss/100; 
            worksheet.getCell("G7").numFmt = "0.00%";
      
      
      //light

          this.actual_per_woven_splice = this.splice_loss_woven * this.ttl_marker_1_4[0].value
         
          this.actual_per_light_splice = this.splice_loss_light * this.ttl_marker_1_4[1].value
         
          this.actual_per_fleece_splice = this.splice_loss_fleece * this.ttl_marker_1_4[2].value
      
          this.actual_per_special_splice = this.splice_loss_special * this.ttl_marker_1_4[3].value
        
          
          this.all_actual_spliceloss = Number(this.actual_per_woven_splice) + Number(this.actual_per_light_splice) + Number(this.actual_per_fleece_splice) + Number(this.actual_per_special_splice)
          
          this.all_actual_spliceloss = this.all_actual_spliceloss / this.ttl_marker_1_4_all[0].value
     




            worksheet.getCell("G8").value = this.all_actual_spliceloss/100;
            worksheet.getCell("G8").numFmt = "0.00%"; 
     
           //cutoutloss

          this.actual_per_woven_cutoutloss = this.cut_loss_woven * this.ttl_marker_1_4[0].value
     
          this.actual_per_light_cutoutloss = this.cut_loss_light * this.ttl_marker_1_4[1].value

          this.actual_per_fleece_cutoutloss = this.cut_loss_fleece * this.ttl_marker_1_4[2].value
    
          this.actual_per_special_cutoutloss = this.cut_loss_special * this.ttl_marker_1_4[3].value
        
          
          this.all_actual_cutoutloss = Number(this.actual_per_woven_cutoutloss) + Number(this.actual_per_light_cutoutloss) + Number(this.actual_per_fleece_cutoutloss) + Number(this.actual_per_special_cutoutloss)
          
          this.all_actual_cutoutloss = this.all_actual_cutoutloss / this.ttl_marker_1_4_all[0].value
   


            worksheet.getCell("G9").value = this.all_actual_cutoutloss/100;
            worksheet.getCell("G9").numFmt = "0.00%";  
     


          //remenants

          this.actual_per_woven_remenants = this.rem_loss_woven * this.ttl_marker_1_4[0].value
     
          this.actual_per_light_remenants = this.rem_loss_light * this.ttl_marker_1_4[1].value

          this.actual_per_fleece_remenants = this.rem_loss_fleece * this.ttl_marker_1_4[2].value
    
          this.actual_per_special_remenants = this.rem_loss_special * this.ttl_marker_1_4[3].value
        
          
          this.all_actual_remnantsloss = Number(this.actual_per_woven_remenants) + Number(this.actual_per_light_remenants) + Number(this.actual_per_fleece_remenants) + Number(this.actual_per_special_remenants)
          
          this.all_actual_remnantsloss = this.all_actual_remnantsloss / this.ttl_marker_1_4_all[0].value



        
            worksheet.getCell("G10").value = this.all_actual_remnantsloss/100; 
            worksheet.getCell("G10").numFmt = "0.00%"; 
      

        this.overall_total = Number(this.all_actual_endloss) + Number(this.all_actual_spliceloss)+ Number(this.all_actual_cutoutloss)
        + Number(this.all_actual_remnantsloss)
      
         worksheet.getCell("G11").value = this.overall_total/100; 
         worksheet.getCell("G11").numFmt = "0.00%"; 

           worksheet.getCell("G14").value = 1.50/100; 
           worksheet.getCell("G14").numFmt = "0.00%"; 


         
          
     
          console.log(this.width_loss_yd_change)
          console.log(this.width_loss_yd_t_change)
      
          for(var ax = 0; ax<this.colume_arr.length; ax++){
          
          if(this.status_week_open == false){
                  if(this.width_loss_yd[ax].value_width_loss   > 0){  
                  worksheet.getCell(this.colume_arr[ax].cn+"13").value = Number(this.width_loss_yd_change[ax].value_width_loss);
                  worksheet.getCell(this.colume_arr[ax].cn+"13").numFmt = "#,##0.00";
                }else{
                  worksheet.getCell(this.colume_arr[ax].cn+"13").value = 0.00;
                  worksheet.getCell(this.colume_arr[ax].cn+"13").numFmt = "0.00";
                }  
          }else{
                if(this.width_loss_yd[ax].value_width_loss   > 0){  
                  worksheet.getCell(this.colume_arr[ax].cn+"13").value = Number(this.width_loss_yd[ax].value_width_loss);
                  worksheet.getCell(this.colume_arr[ax].cn+"13").numFmt = "#,##0.00";
                }else{
                  worksheet.getCell(this.colume_arr[ax].cn+"13").value = 0.00;
                  worksheet.getCell(this.colume_arr[ax].cn+"13").numFmt = "0.00";
                }  
          }
                
         } 


          for(var ax = 0; ax<this.colume_arr.length; ax++){
           this.ac_width2 = 0
           if(this.status_week_open == false){
             this.ac_width2 = this.width_loss_yd_change[ax].value_width_loss / this.ttl_marker_yd_change[ax].value_ttl *100
           }else{
           this.ac_width2 = this.width_loss_yd[ax].value_width_loss / this.ttl_marker_yd[ax].value_ttl *100
           }
         
            if(this.ac_width2  > 0){
                  worksheet.getCell(this.colume_arr[ax].cn+"14").value = this.ac_width2/100;
                  worksheet.getCell(this.colume_arr[ax].cn+"14").numFmt = "0.00%";
                }else{
                  worksheet.getCell(this.colume_arr[ax].cn+"14").value = 0.00;
                  worksheet.getCell(this.colume_arr[ax].cn+"14").numFmt = "0.00%";
                }      
         } 

          this.sum_all_week_yard_total = 0
            this.commulate_dozen_cut_total = 0

            if(this.count_round_week == 1){
     if(this.status_week_open == false){
          for(var ax = 0; ax<this.all_week_cut_yard_4_4_1.length; ax++){
            if(ax == 0 || ax == 1 || ax == 2){
            this.all_week_cut_yard_4_4_1_change.push({
              value_cut:""
            })
            }else{
              this.all_week_cut_yard_4_4_1_change.push({
              value_cut:this.all_week_cut_yard_4_4_1[ax].value_cut
            })
            }
          }
     }
     if(this.status_week_open == false){
          for(var ax = 0; ax<this.all_week_cut4_4.length; ax++){
            if(ax == 0 || ax == 1 || ax == 2){
            this.all_week_cut4_4_change.push({
              value_cut:""
            })
            }else{
              this.all_week_cut4_4_change.push({
              value_cut:this.all_week_cut4_4[ax].value_cut
            })
            }
          }
     }
     
    }else if(this.count_round_week == 2){
        if(this.status_week_open == false){
          for(var ax = 0; ax<this.all_week_cut_yard_4_4_1.length; ax++){
            if(ax == 0 || ax == 1){
            this.all_week_cut_yard_4_4_1_change.push({
              value_cut:""
            })
            }else{
              this.all_week_cut_yard_4_4_1_change.push({
              value_cut:this.all_week_cut_yard_4_4_1[ax].value_cut
            })
            }
          }
     }

     if(this.status_week_open == false){
          for(var ax = 0; ax<this.all_week_cut4_4.length; ax++){
            if(ax == 0 || ax == 1){
            this.all_week_cut4_4_change.push({
              value_cut:""
            })
            }else{
              this.all_week_cut4_4_change.push({
              value_cut:this.all_week_cut4_4[ax].value_cut
            })
            }
          }
     }
    }else{
      if(this.status_week_open == false){
          for(var ax = 0; ax<this.all_week_cut_yard_4_4_1.length; ax++){
            if(ax == 0){
            this.all_week_cut_yard_4_4_1_change.push({
              value_cut:""
            })
            }else{
              this.all_week_cut_yard_4_4_1_change.push({
              value_cut:this.all_week_cut_yard_4_4_1[ax].value_cut
            })
            }
          }
     }
     if(this.status_week_open == false){
          for(var ax = 0; ax<this.all_week_cut4_4.length; ax++){
            if(ax == 0 || ax == 1){
            this.all_week_cut4_4_change.push({
              value_cut:""
            })
            }else{
              this.all_week_cut4_4_change.push({
              value_cut:this.all_week_cut4_4[ax].value_cut
            })
            }
          }
     }
    }
     
         for(var ax = 0; ax<this.colume_arr.length; ax++){
           this.sum_all_week_yard = 0
          if(this.status_week_open == false){
             this.sum_all_week_yard = Number(this.all_week_cut_yard_4_4_1_change[ax].value_cut) + Number(this.all_week_cut_yard_4_4_1_change[ax].value_cut)
          }else{
               this.sum_all_week_yard = Number(this.all_week_cut_yard_4_4_1[ax].value_cut) + Number(this.all_week_cut_yard_4_4_1[ax].value_cut)
          }
         
            if(this.sum_all_week_yard   > 0){
                  worksheet.getCell(this.colume_arr[ax].cn+"15").value = Number(this.sum_all_week_yard);
                  worksheet.getCell(this.colume_arr[ax].cn+"15").numFmt = "#,##0.00";
                }else{
                  worksheet.getCell(this.colume_arr[ax].cn+"15").value = 0.00;
                  worksheet.getCell(this.colume_arr[ax].cn+"15").numFmt = "0.00";
                }
            
            this.sum_all_week_yard_total = Number(this.sum_all_week_yard_total) + Number(this.sum_all_week_yard)
              
             this.audit_per_total = 0

            if(this.status_week_open == false){
              this.audit_per_total =  this.ttl_marker_yd_change[ax].value_ttl/this.sum_all_week_yard *100
            }else{
              this.audit_per_total =  this.ttl_marker_yd[ax].value_ttl/this.sum_all_week_yard *100
            }
            
            
             if(this.audit_per_total  > 0){
                  worksheet.getCell(this.colume_arr[ax].cn+"16").value = this.audit_per_total/100
                  worksheet.getCell(this.colume_arr[ax].cn+"16").numFmt = "0.00%";
                }else{
                  worksheet.getCell(this.colume_arr[ax].cn+"16").value = 0.00;
                  worksheet.getCell(this.colume_arr[ax].cn+"16").numFmt = "0.00%";
                }
          

              this.dozen_cut_total = 0
            if(this.status_week_open == false){
              this.dozen_cut_total =  this.all_week_cut4_4_change[ax].value_cut/12
            }else{
              this.dozen_cut_total =  this.all_week_cut4_4[ax].value_cut/12
            }
            
             if(this.dozen_cut_total  > 0){
                  worksheet.getCell(this.colume_arr[ax].cn+"17").value = Number(this.dozen_cut_total);
                  worksheet.getCell(this.colume_arr[ax].cn+"17").numFmt = "#,##0.00";
                }else{
                  worksheet.getCell(this.colume_arr[ax].cn+"17").value = 0.00;
                  worksheet.getCell(this.colume_arr[ax].cn+"17").numFmt = "0.00";
                } 
              
                this.commulate_dozen_cut_total = Number(this.commulate_dozen_cut_total) + Number(this.dozen_cut_total)
                
         } 
        
         worksheet.getCell("M15").value = this.sum_all_week_yard_total;
        worksheet.getCell("M15").numFmt = "#,##0.00";
        

          this.total_audit_fabric_cut = 0
        this.total_audit_fabric_cut = this.sum_all_ttl_marker_4_week / this.sum_all_week_yard_total *100

        worksheet.getCell("L16").value = this.total_audit_fabric_cut/100;
        worksheet.getCell("L16").numFmt = "0.00%";
        
        worksheet.getCell("M17").value = Number(this.commulate_dozen_cut_total);
        worksheet.getCell("M17").numFmt = "#,##0.00";

        this.all_total_dozen_cut = 0
        for(var ax = 0; ax<this.all_week_yard.length; ax++){
          this.all_total_dozen_cut = Number(this.all_total_dozen_cut) + Number(this.all_week_yard[ax].value_cut_yard)
          
        }
        if(this.all_total_dozen_cut > 0){
       
          this.all_total_dozen_cut = this.all_total_dozen_cut/12
        worksheet.getCell("O17").value = Number(this.all_total_dozen_cut);
        worksheet.getCell("O17").numFmt = "#,##0.00";
        }else{
          worksheet.getCell("O17").value = 0.00;
          worksheet.getCell("O17").numFmt = "0.00";
        }
        //ข้อมูลด้านขวาล่าง
        
       worksheet.getCell("R18").value = "% Width Loss Target";
        for(var ax = 0; ax<this.date_use_data2.length; ax++){
          worksheet.getCell(this.arr_colume2[ax].name_col+"18").value = 1.50/100;
          worksheet.getCell(this.arr_colume2[ax].name_col+"18").numFmt = "0.00%";
          worksheet.getCell(this.arr_colume2[ax].name_col+"19").value = "";
        }
        worksheet.getCell("R19").value = "% STD. Width Loss";
        

        worksheet.getCell("R20").value = "Average Act. Width";
        worksheet.getCell("R21").value = "% Acc Act. Width Loss";


      /*   for(var j = 0; j<this.all_data_arr.length; j++){
           if(this.all_data_arr[j].obs_width !== ""){
        this.total_width_inch = this.all_data_arr[j].obs_width * this.all_data_arr[j].ttl_marker;
        this.add_width_inch = Number(this.add_width_inch) + Number(this.total_width_inch);
           }
        }

       
        
        if(this.result_ttlmarker !== 0){
        this.add_width_inch = this.add_width_inch / this.result_ttlmarker;
        } */
        this.all_val_ttl = 0
       for(var ax = 0; ax<this.date_use_data2.length; ax++){
         this.all_val_ttl = Number(this.all_val_ttl) + Number(this.ttl_marker_yd2[ax].value_ttl)
       }
  
        for(var ax = 0; ax<this.date_use_data2.length; ax++){
          
             worksheet.getCell(this.arr_colume2[ax].name_col+"20").value = Number(this.avg_width_loss_cal[ax].value);
             worksheet.getCell(this.arr_colume2[ax].name_col+"20").numFmt = '0.00'; 
        }


       

/* 
               this.total_spread_width_loss.push({
                value_width_loss : this.total_spread_loss
              })    */

          this.sum_value_ttl_marker = 0
          this.sum_value_width_loss = 0
          
          this.last_result_acc_act_width_loss = 0

       

        for(var ax = 0; ax<this.date_use_data2.length; ax++){

          
          if(isNaN(this.total_spread_width_loss[ax].value_width_loss )== true){
            this.total_spread_width_loss[ax].value_width_loss = 0
          }
          if(isNaN(this.ttl_marker_yd2[ax].value_ttl )== true){
            this.ttl_marker_yd2[ax].value_ttl = 0
          }this.width_loss_yd_t[ax].value_width_loss
        
        this.sum_value_width_loss = Number(this.sum_value_width_loss) + Number(this.total_spread_width_loss[ax].value_width_loss)
        this.sum_value_ttl_marker = Number(this.sum_value_ttl_marker) + Number(this.ttl_marker_yd2[ax].value_ttl)
        
        this.last_result_acc_act_width_loss = this.sum_value_width_loss / this.sum_value_ttl_marker *100

         
        
           if(this.last_result_acc_act_width_loss > 0){
          worksheet.getCell(this.arr_colume2[ax].name_col+"21").value = this.last_result_acc_act_width_loss/100;
          worksheet.getCell(this.arr_colume2[ax].name_col+"21").numFmt = '0.00%'; 
          }else{
          worksheet.getCell(this.arr_colume2[ax].name_col+"21").value = 0.00;
          worksheet.getCell(this.arr_colume2[ax].name_col+"21").numFmt = '0.00%'; 
          }
        }

           worksheet.getCell("T24").value = "Diff";
           worksheet.getCell("U24").value = "Actual";        
           worksheet.getCell("V24").value = "Target";
           worksheet.getCell("V25").value = this.total_loss[0].p/100;
           worksheet.getCell("V25").numFmt = '0.00%'
           worksheet.getCell("W25").value = "Type1(Woven)";
           worksheet.getCell("V26").value = this.total_loss[1].p/100;
           worksheet.getCell("V26").numFmt = '0.00%'
           worksheet.getCell("W26").value = "Type2(L&Mid.Wt.)";
           worksheet.getCell("V27").value = this.total_loss[2].p/100;
           worksheet.getCell("V27").numFmt = '0.00%'
           worksheet.getCell("W27").value = "Type3(Fleece)";
           worksheet.getCell("V28").value = this.total_loss[3].p/100;
           worksheet.getCell("V28").numFmt = '0.00%'
           worksheet.getCell("W28").value = "Type4(Special)";

        /* this.actual_knit_woven = 0   //M7
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
        } */
      
      this.all_ttl_woven = 0
      this.all_ttl_light = 0
      this.all_ttl_fleece = 0
      this.all_ttl_special = 0
      
      this.sum_keep_end_woven = 0
      this.sum_keep_end_light = 0
      this.sum_keep_end_fleece = 0
      this.sum_keep_end_special = 0


      this.sum_keep_splice_woven = 0
      this.sum_keep_splice_light = 0
      this.sum_keep_splice_fleece = 0
      this.sum_keep_splice_special = 0


      this.sum_keep_cut_woven = 0
      this.sum_keep_cut_light = 0
      this.sum_keep_cut_fleece = 0
      this.sum_keep_cut_special = 0


      this.sum_keep_rem_woven = 0
      this.sum_keep_rem_light = 0
      this.sum_keep_rem_fleece = 0
      this.sum_keep_rem_special = 0

   
        
        this.sum_keep_end_woven = Number(this.sum_keep_end_woven) + Number(this.keep_end_woven[0].value)
        this.sum_keep_end_light = Number(this.sum_keep_end_light) + Number(this.keep_end_light[0].value)
        this.sum_keep_end_fleece = Number(this.sum_keep_end_fleece) + Number(this.keep_end_fleece[0].value)
        this.sum_keep_end_special = Number(this.sum_keep_end_special) + Number(this.keep_end_special[0].value)

        this.sum_keep_splice_woven = Number(this.sum_keep_splice_woven) + Number(this.keep_splice_woven[0].value)
        this.sum_keep_splice_light = Number(this.sum_keep_splice_light) + Number(this.keep_splice_light[0].value)
        this.sum_keep_splice_fleece = Number(this.sum_keep_splice_fleece) + Number(this.keep_splice_fleece[0].value)
        this.sum_keep_splice_special = Number(this.sum_keep_splice_special) + Number(this.keep_splice_special[0].value)

        this.sum_keep_cut_woven = Number(this.sum_keep_cut_woven) + Number(this.keep_cut_woven[0].value)
        this.sum_keep_cut_light = Number(this.sum_keep_cut_light) + Number(this.keep_cut_light[0].value)
        this.sum_keep_cut_fleece = Number(this.sum_keep_cut_fleece) + Number(this.keep_cut_fleece[0].value)
        this.sum_keep_cut_special = Number(this.sum_keep_cut_special) + Number(this.keep_cut_special[0].value)

        this.sum_keep_rem_woven = Number(this.sum_keep_rem_woven) + Number(this.keep_rem_woven[0].value)
        this.sum_keep_rem_light = Number(this.sum_keep_rem_light) + Number(this.keep_rem_light[0].value)
        this.sum_keep_rem_fleece = Number(this.sum_keep_rem_fleece) + Number(this.keep_rem_fleece[0].value)
        this.sum_keep_rem_special = Number(this.sum_keep_rem_special) + Number(this.keep_rem_special[0].value)
     // }
      
      
   
      this.actual_end_woven = 0
      this.actual_end_light = 0
      this.actual_end_fleece = 0
      this.actual_end_special = 0


      this.actual_splice_woven = 0
     
      this.actual_splice_fleece = 0
      this.actual_splice_special = 0

      this.actual_cut_woven = 0
      this.actual_cut_light = 0
      this.actual_cut_fleece = 0
      this.actual_cut_special = 0

      this.actual_rem_woven = 0
      this.actual_rem_light = 0
      this.actual_rem_fleece = 0
      this.actual_rem_special = 0
      
   

    

       this.actual_end_woven = this.sum_keep_end_woven / this.all_ttl_marker1_4_diff[0].value_ttl *100
       this.actual_end_light = this.sum_keep_end_light / this.all_ttl_marker1_4_diff[1].value_ttl *100
       this.actual_end_fleece = this.sum_keep_end_fleece / this.all_ttl_marker1_4_diff[2].value_ttl *100
       this.actual_end_special = this.sum_keep_end_special / this.all_ttl_marker1_4_diff[3].value_ttl *100

     
       this.actual_splice_woven = this.sum_keep_splice_woven / this.all_ttl_marker1_4_diff[0].value_ttl *100
       this.actual_splice_light = this.sum_keep_splice_light / this.all_ttl_marker1_4_diff[1].value_ttl *100
       this.actual_splice_fleece = this.sum_keep_splice_fleece / this.all_ttl_marker1_4_diff[2].value_ttl *100
       this.actual_splice_special = this.sum_keep_splice_special / this.all_ttl_marker1_4_diff[3].value_ttl *100


   

       this.actual_cut_woven = this.sum_keep_cut_woven / this.all_ttl_marker1_4_diff[0].value_ttl *100
       this.actual_cut_light = this.sum_keep_cut_light / this.all_ttl_marker1_4_diff[1].value_ttl *100
       this.actual_cut_fleece = this.sum_keep_cut_fleece / this.all_ttl_marker1_4_diff[2].value_ttl *100
       this.actual_cut_special = this.sum_keep_cut_special / this.all_ttl_marker1_4_diff[3].value_ttl *100

  



       this.actual_rem_woven = this.sum_keep_rem_woven / this.all_ttl_marker1_4_diff[0].value_ttl *100
       this.actual_rem_light = this.sum_keep_rem_light / this.all_ttl_marker1_4_diff[1].value_ttl *100
       this.actual_rem_fleece = this.sum_keep_rem_fleece / this.all_ttl_marker1_4_diff[2].value_ttl *100
       this.actual_rem_special = this.sum_keep_rem_special / this.all_ttl_marker1_4_diff[3].value_ttl *100
       



       this.total_actual_knit_woven = 0

       this.total_actual_knit_light = 0

       this.total_actual_knit_fleece = 0

       this.total_actual_knit_special = 0

       this.total_actual_knit_woven = Number(this.actual_end_woven ) + Number(this.actual_splice_woven) + Number(this.actual_cut_woven) + Number(this.actual_rem_woven)

       this.total_actual_knit_light = Number(this.actual_end_light) + Number(this.actual_splice_light) + Number(this.actual_cut_light) + Number(this.actual_rem_light)

       this.total_actual_knit_fleece = Number(this.actual_end_fleece ) + Number(this.actual_splice_fleece) + Number(this.actual_cut_fleece) + Number(this.actual_rem_fleece)

       this.total_actual_knit_special = Number(this.actual_end_special ) + Number(this.actual_splice_special) + Number(this.actual_cut_special) + Number(this.actual_rem_special)

      if(this.total_actual_knit_woven > 0){
      worksheet.getCell("U25").value = this.total_actual_knit_woven/100
      worksheet.getCell("U25").numFmt = '0.00%'; 
      }else{
        worksheet.getCell("U25").value = 0.00
        worksheet.getCell("U25").numFmt = '0.00%'; 
      }

      if(this.total_actual_knit_light > 0){
      worksheet.getCell("U26").value = this.total_actual_knit_light/100
      worksheet.getCell("U26").numFmt = '0.00%'; 
      }else{
        worksheet.getCell("U26").value = 0.00
        worksheet.getCell("U26").numFmt = '0.00%'; 
      }

      if(this.total_actual_knit_fleece > 0){
      worksheet.getCell("U27").value = this.total_actual_knit_fleece/100
      worksheet.getCell("U27").numFmt = '0.00%'; 
      }else{
        worksheet.getCell("U27").value = 0.00
        worksheet.getCell("U27").numFmt = '0.00%'; 
      }

      if(this.total_actual_knit_special > 0){
      worksheet.getCell("U28").value = this.total_actual_knit_special/100
      worksheet.getCell("U28").numFmt = '0.00%'; 
      }else{
        worksheet.getCell("U28").value = 0.00
        worksheet.getCell("U28").numFmt = '0.00%'; 
      }

      this.diff_woven = 0
      this.diff_light = 0
      this.diff_fleece = 0
      this.diff_special = 0

      if(this.total_actual_knit_woven < 0){
        this.total_actual_knit_woven = 0
      }

      if(this.total_actual_knit_light < 0){
        this.total_actual_knit_light = 0
      }

      if(this.total_actual_knit_fleece < 0){
        this.total_actual_knit_fleece = 0
      }

      if(this.total_actual_knit_special < 0){
        this.total_actual_knit_special = 0
      }
      this.diff_woven = Number(this.total_actual_knit_woven) - Number(this.total_loss[0].p)

      this.diff_light = Number(this.total_actual_knit_light) - Number(this.total_loss[1].p)

      this.diff_fleece = Number(this.total_actual_knit_fleece) - Number(this.total_loss[2].p)

      this.diff_special = Number(this.total_actual_knit_special) - Number(this.total_loss[3].p)
      
      if(isNaN(this.total_actual_knit_woven)==true){
        this.total_actual_knit_woven = 0
      }
      if(isNaN(this.diff_woven)==true){
        this.diff_woven = 0
      }
      this.result_diff.push({
        name:"Type1(Woven)",
        target:"0.38%",
        actual:this.total_actual_knit_woven,
        diff:this.diff_woven
      })

      if(isNaN(this.total_actual_knit_light)==true){
        this.total_actual_knit_light = 0
      }
      if(isNaN(this.diff_light)==true){
        this.diff_light = 0
      }
      this.result_diff.push({
        name:"Type2(L&Mid.Wt.)",
        target:"0.50%",
        actual:this.total_actual_knit_light,
        diff:this.diff_light
      })

      if(isNaN(this.total_actual_knit_fleece)==true){
        this.total_actual_knit_fleece = 0
      }
      if(isNaN(this.diff_fleece)==true){
        this.diff_fleece = 0
      }

      this.result_diff.push({
        name:"Type3(Fleece)",
        target:"0.56%",
        actual:this.total_actual_knit_fleece,
        diff:this.diff_fleece
      })

      if(isNaN(this.total_actual_knit_special)==true){
        this.total_actual_knit_special = 0
      }
      if(isNaN(this.diff_special)==true){
        this.diff_special = 0
      }

      this.result_diff.push({
        name:"Type4(Special)",
        target:"0.65%",
        actual:this.total_actual_knit_special,
        diff:this.diff_special
      })

      if(isNaN(this.diff_woven) == false){
      worksheet.getCell("T25").value = this.diff_woven/100
      worksheet.getCell("T25").numFmt = '0.00%'; 
      }else{
      worksheet.getCell("T25").value = 0.00
      worksheet.getCell("T25").numFmt = '0.00%'; 
      }

      if(isNaN(this.diff_light) == false){
      worksheet.getCell("T26").value = this.diff_light/100
      worksheet.getCell("T26").numFmt = '0.00%'; 
      }else{
      worksheet.getCell("T26").value = 0.00
      worksheet.getCell("T26").numFmt = '0.00%'; 
      }

      if(isNaN(this.diff_fleece) == false){
      worksheet.getCell("T27").value = this.diff_fleece/100
      worksheet.getCell("T27").numFmt = '0.00%'; 
      }else{
        worksheet.getCell("T27").value = 0.00
        worksheet.getCell("T27").numFmt = '0.00%'; 
      }
      if(isNaN(this.diff_special) == false){
      worksheet.getCell("T28").value = this.diff_special/100
      worksheet.getCell("T28").numFmt = '0.00%'; 
      }else{
        worksheet.getCell("T28").value = 0.00
        worksheet.getCell("T28").numFmt = '0.00%'; 
      }

       worksheet.getCell("T32").value = "Detail"
       worksheet.getCell("T33").value = "End Loss"
       worksheet.getCell("T34").value = "Splice Loss"
       worksheet.getCell("T35").value = "Cut out Loss"
       worksheet.getCell("T36").value = "Remnants Loss"
       worksheet.getCell("T37").value = "Total Length Loss"
       worksheet.getCell("T38").value = "Width Loss"

       worksheet.getCell("U32").value = "Target"
       worksheet.getCell("V32").value = "Actual"
       worksheet.getCell("W32").value = "Diff"
       worksheet.getCell("X32").value = "Cummalated since 25-10-2021"
      
      worksheet.getCell("U33").value =  this.all_actual_endloss/100
      worksheet.getCell("U33").numFmt = '0.00%';
      worksheet.getCell("V33").value =  this.endloss_ac/100
      worksheet.getCell("V33").numFmt = '0.00%';
      this.diff_actual_end = Number(this.endloss_ac) - Number(this.all_actual_endloss)
      worksheet.getCell("W33").value =  this.diff_actual_end /100
      worksheet.getCell("W33").numFmt = '0.00%';
   

       this.result_actual.push({
        name:"End Loss",
        actual:this.endloss_ac
      })

      worksheet.getCell("U34").value =  this.all_actual_spliceloss/100
      worksheet.getCell("U34").numFmt = '0.00%';
      worksheet.getCell("V34").value =  this.splice_ac/100
      worksheet.getCell("V34").numFmt = '0.00%';
      this.diff_actual_splice = Number(this.splice_ac) - Number(this.all_actual_spliceloss)
      worksheet.getCell("W34").value =  this.diff_actual_splice /100
      worksheet.getCell("W34").numFmt = '0.00%';

     

       this.result_actual.push({
        name:"Splice Loss",
        actual:this.splice_ac
      })


      worksheet.getCell("U35").value =  this.all_actual_cutoutloss/100
      worksheet.getCell("U35").numFmt = '0.00%';
      worksheet.getCell("V35").value =  this.cut_ac/100
      worksheet.getCell("V35").numFmt = '0.00%';
      this.diff_actual_cut = Number(this.cut_ac) - Number(this.all_actual_cutoutloss)
      worksheet.getCell("W35").value =  this.diff_actual_cut /100
      worksheet.getCell("W35").numFmt = '0.00%';


       this.result_actual.push({
        name:"Cut out Loss",
        actual:this.cut_ac
      })



      worksheet.getCell("U36").value =  this.all_actual_remnantsloss/100
      worksheet.getCell("U36").numFmt = '0.00%';
      worksheet.getCell("V36").value =  this.rem_ac/100
      worksheet.getCell("V36").numFmt = '0.00%';
      this.diff_actual_rem = Number(this.rem_ac) - Number(this.all_actual_remnantsloss)
      worksheet.getCell("W36").value =  this.diff_actual_rem /100
      worksheet.getCell("W36").numFmt = '0.00%';

 
       this.result_actual.push({
        name:"Remnants Loss",
        actual:this.rem_ac
      })

//this.overall_total
      worksheet.getCell("U37").value =  this.overall_total/100
      worksheet.getCell("U37").numFmt = '0.00%';
      worksheet.getCell("V37").value =  this.total_ac/100
      worksheet.getCell("V37").numFmt = '0.00%';
      this.diff_actual_total = Number(this.total_ac) - Number(this.overall_total)
      worksheet.getCell("W37").value =  this.diff_actual_total /100
      worksheet.getCell("W37").numFmt = '0.00%'; 

//width loss
      worksheet.getCell("U38").value =  1.50/100
      worksheet.getCell("U38").numFmt = '0.00%';
      this.ac_width_v = 0
      this.ac_width_v = this.width_loss_yd[3].value_width_loss / this.ttl_marker_yd[3].value_ttl *100
      worksheet.getCell("V38").value =  this.ac_width_v/100
      worksheet.getCell("V38").numFmt = '0.00%';
      this.diff_actual_widthloss = Number(this.ac_width_v) - Number(1.50)
      worksheet.getCell("W38").value =  this.diff_actual_widthloss/100
      worksheet.getCell("W38").numFmt = '0.00%'; 

      
     this.all_ttl_commulate_widthloss = 0
     this.all_ttl_commulate_widthloss = Number(this.commulate_end) + Number(this.commulate_splice) + Number(this.commulate_cut)
     +Number(this.commulate_rem)

      worksheet.getCell("X33").value =  this.commulate_end /100
      worksheet.getCell("X33").numFmt = '0.00%';
      worksheet.getCell("X34").value =  this.commulate_splice /100
      worksheet.getCell("X34").numFmt = '0.00%';
      worksheet.getCell("X35").value =  this.commulate_cut /100
      worksheet.getCell("X35").numFmt = '0.00%';
      worksheet.getCell("X36").value =  this.commulate_rem /100
      worksheet.getCell("X36").numFmt = '0.00%';
      worksheet.getCell("X37").value =  this.all_ttl_commulate_widthloss /100
      worksheet.getCell("X37").numFmt = '0.00%';
      worksheet.getCell("X38").value =  this.total_spread_loss_per/100
      worksheet.getCell("X38").numFmt = '0.00%';


      console.log(this.result_diff)
        this.find_max_diff = 0
        this.find_max_name =""
        this.find_max_target =""
        this.find_max_actual = ""
        this.count5 = 0
       // ค่า value
      //debugger
        for(ax = 0; ax<this.result_diff.length; ax++){
        if(this.count5 == 0){
          if(this.result_diff[ax].diff !== 0){
            this.find_max_diff  = this.result_diff[ax].diff
            this.find_max_name = this.result_diff[ax].name
            this.find_max_target = this.result_diff[ax].target
            this.find_max_actual = this.result_diff[ax].actual
          this.count5 ++
         }  
         }else{
           if(this.result_diff[ax].diff !== 0){
             if(this.find_max_diff < this.result_diff[ax].diff){
                this.find_max_diff  = this.result_diff[ax].diff
            this.find_max_name = this.result_diff[ax].name
            this.find_max_target = this.result_diff[ax].target
            this.find_max_actual = this.result_diff[ax].actual
             }

           }
           this.count5 ++
         }
        }
         
        
        console.log(this.find_max_diff)
        console.log(this.find_max_name)
        console.log(this.find_max_target)
        console.log(this.find_max_actual)


  //หาค่าHighest Loss
      
        this.find_max_name2 =""
     
        this.find_max_value2 = ""
        this.count6 = 0
       // ค่า value
       
        

        this.minus_end = 0
        this.minus_splice = 0
        this.minus_cut = 0 
        this.minus_rem = 0

       

        this.minus_end = this.end_loss_4_week - this.all_actual_endloss
        this.all_minus_value.push({
          value:this.minus_end,
          value_name:"End Loss"
        })
        this.minus_splice =  this.splice_loss_4_week - this.all_actual_spliceloss
        this.all_minus_value.push({
          value:this.minus_splice,
          value_name:"Splice Loss"
        })
        this.minus_cut =  this.cut_loss_4_week - this.all_actual_cutoutloss
        this.all_minus_value.push({
          value:this.minus_cut,
          value_name:"Cut Out Loss"
        })
        this.minus_rem =  this.rem_loss_4_week - this.all_actual_remnantsloss
        this.all_minus_value.push({
          value:this.minus_rem,
          value_name:"Remnants Loss"
        })

        
        //debugger
       for(ax = 0; ax<this.all_minus_value.length; ax++){
          if(this.count6 == 0){
            this.find_max_name2 = this.all_minus_value[ax].value_name
            this.find_max_value2 = this.all_minus_value[ax].value
             this.count6 ++
         }else{
           if(this.find_max_value2  == 0){
            this.find_max_name2 = this.all_minus_value[ax].value_name
            this.find_max_value2 = this.all_minus_value[ax].value
           }else if( this.find_max_value2  < this.all_minus_value[ax].value){
            this.find_max_name2 = this.all_minus_value[ax].value_name
            this.find_max_value2 = this.all_minus_value[ax].value
           }else if(this.find_max_value2  == this.all_minus_value[ax].value){
            this.find_max_name2 = this.all_minus_value[ax].value_name
            this.find_max_value2 = this.all_minus_value[ax].value
            this.count6 ++
           }
           
         }
         
        }  

      
                                                                                                                   
       worksheet.getCell("M19").value = "สรุปผลการดำเนินการสัปดาห์ปัจจุบัน\n%Total Spread Loss Target ="+parseFloat(this.overall_total).toFixed(2)+"%"+"\n Actual ="+parseFloat(this.total_ac).toFixed(2)+"%"+"\nDiff = "+parseFloat(this.diff_actual_total).toFixed(2)+"%" +"\nHighest Loss = "+this.find_max_name2+"\nFabric Type ที่ส่งผลให้Lossสูงมาจาก = "+this.find_max_name+"\nTarget="+parseFloat(this.find_max_target).toFixed(2)+"%"+ "\nActual ="+parseFloat(this.find_max_actual).toFixed(2)+"%"+"\nDiff ="+parseFloat(this.find_max_diff).toFixed(2)+"%";
     
       worksheet.mergeCells("M19:O36");
       
          worksheet.getCell("M19").font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
        worksheet.getCell("M19").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
       worksheet.getCell("M19").alignment = {
          vertical: "top",
          horizontal: "center",
          wrapText: true
        };

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

        worksheet.mergeCells('B19:L34')
        worksheet.getCell('B19').fill={
          type: 'pattern',
          pattern:'none',
          fgColor:{argb:'FFFFFFFF'},
          bgColor:{argb:'FFFFFFFF'},
        }; 
      this.$q.loading.hide({
          
        })
      
      
     
        workbook.xlsx.writeBuffer().then((data) => {
        const blob = new Blob([data], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8",
        });
        saveAs(blob, "Report Weekly SLA3-1.xlsx"); 
      });    
      for(var ax=0; ax<52; ax++){
          this.total_ttl_marker1_4 = 0
          this.total_ttl_marker1_4 = Number(this.ttl_marker_endless_woven[ax].value) + Number(this.ttl_marker_endless_light[ax].value)
          +Number(this.ttl_marker_endless_fleece[ax].value) + Number(this.ttl_marker_endless_special[ax].value)
           this.total_target.push({
             all_ttl : this.total_ttl_marker1_4
           })
      }

     
      
      for(var ax = 0; ax<52; ax++){
          this.total_spread_target = 0
          this.total_spread_target = (Number(this.phase_type[0].per * this.ttl_marker_endless_woven[ax].value)  + Number(this.phase_type[1].per * this.ttl_marker_endless_light[ax].value)
          +Number(this.phase_type[2].per * this.ttl_marker_endless_fleece[ax].value) + Number(this.phase_type[3].per *  this.ttl_marker_endless_special[ax].value)) /this.total_target[ax].all_ttl 
          if(isNaN(this.total_spread_target)==true){
            this.total_spread_target = 0
          }
          if(this.total_spread_target > 0){
          this.ttl_target.push({
            value:this.total_spread_target,
            value_ax:ax,
          })
          }
      }


      for(var ax = 0; ax<this.ttl_target.length; ax++){
        
         this.date_data_for_graph.push({
           value:this.date_use_data2[this.ttl_target[ax].value_ax].first_date+" - "+ this.date_use_data2[this.ttl_target[ax].value_ax].sec_date
         })

         this.ttl_spread_graph_last.push({
           value:this.ttl_spread_graph[this.ttl_target[ax].value_ax].value
         })
        
      } 
     

     let target_ttl = []
      let index2 = 0;
      this.ttl_target.forEach(e => {
                target_ttl[index2] = parseFloat(e.value).toFixed(2)
                index2++
            })
        


      let data_ttl = []
      let index1 = 0;
      this.ttl_spread_graph_last .forEach(e => {
                data_ttl[index1] = parseFloat(e.value).toFixed(2)
                index1++
            })
          
      
      let data_date = [];
            let index = 0;
            this.date_data_for_graph .forEach(e => {
                data_date[index] = e.value
                index++
            })
        
console.log(data_ttl)
console.log(target_ttl)
console.log(data_date)



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
    boundaryGap: [1, '200%']
  },
  
  series: [
    {
      name : 'actual',
      data: data_ttl,
      type: 'line'
    },
    {
      name: 'target',
      data: target_ttl,
      type: 'line'
    }
  ]
};

option && myChart.setOption(option);



     }  

     /*  }catch(err){
         this.$q.loading.hide({
          
        })
       this.$q.notify({
          message: "Something is Wrong",
          color: "red-9",
        });
        console.log(err)
      }  */
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