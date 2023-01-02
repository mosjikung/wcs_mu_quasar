<template>
  <div class="row justify-center">
    <q-input class="q-px-xl" filled v-model="start">
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
     <q-input class="q-px-xl" filled v-model="end">
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
      class="q-px-xl"
      v-model="org"
      :options="option_org"
      style="width: 200px"
      disable
      label="Chose Org"
    />

    <q-btn
      size="md"
      dense
      class="q-px-xl q-py-xl"
      color="primary"
      label="Export"
      @click="exportexcel()"
    />
    <br>
  <div class="col-12">
   <div class="row justify-center">  
      <q-input class="q-px-xl q-py-md" filled v-model="start_h">
      <template v-slot:append>
        <q-icon name="event" class="cursor-pointer">
          <q-popup-proxy
            ref="qDateProxy"
            transition-show="scale"
            transition-hide="scale"
          >
            <q-date v-model="start_h" mask="DD/MM/YYYY">
              <div class="row items-center justify-end">
                <q-btn v-close-popup label="Close" color="primary" flat />
              </div>
            </q-date>
          </q-popup-proxy>
        </q-icon>
      </template>
    </q-input>
     <q-input class="q-px-xl q-py-md" filled v-model="end_h">
      <template v-slot:append>
        <q-icon name="event" class="cursor-pointer">
          <q-popup-proxy
            ref="qDateProxy"
            transition-show="scale"
            transition-hide="scale"
          >
            <q-date v-model="end_h" mask="DD/MM/YYYY">
              <div class="row items-center justify-end">
                <q-btn v-close-popup label="Close" color="primary" flat />
              </div>
            </q-date>
          </q-popup-proxy>
        </q-icon>
      </template>
    </q-input>

    <div class="q-py-lg q-px-md">
        <p class="text-danger">***Date For Spread Loss By floor-h***</p>
    </div>

  


   </div>
   </div>

  </div>

 
</template>
<script>
import { ref } from "vue";
import { date } from "quasar";
import axios from "axios";
import { useQuasar, QSpinnerGears } from "quasar";
import { onBeforeUnmount } from "vue";
import * as Excel from "exceljs";
import { saveAs } from "file-saver";
import moment from "moment";
import * as echarts from 'echarts';
export default {
    data(){
        return{
            org:"",
            start:"",
            chose_graph:"",
            end:"",
            start_date:"",
            end_date:"",
            start_h:"",
            end_h:"",
            date_use_data:[],
            date_use_data2:[],
            fix_line:20,
            rowexport:[],
            rowexport2:[],
            rowexport_type_1:[],
            rowexport_type_2:[],
            rowexport_type_3:[],
            rowexport_type_4:[],
            rowexport_type_5:[],
            rowexport_type_6:[],
            rowexport_type_7:[],
            rowexport_type_8:[],
            rowexport_type_9:[],
            rowexport_type_10:[],
            rowexport_type_11:[],
            rowexport_type_12:[],
            rowexport_type_13:[],
            rowexport_type_14:[],
            rowexport_type_15:[],
            rowexport_type_16:[],
            rowexport_type_17:[],
            rowexport_type_18:[],
            rowexport_type_19:[],
            rowexport_type_20:[],
            rowexport_type_M1:[],
            rowexport_type_M2:[],
            rowexport_type_M3:[],
            rowexport_type_M4:[],
            rowexport_type_M5:[],
            rowexport_type_M6:[],
            rowexport_type_M7:[],
            rowexport_type_M8:[],
            rowexport_type_M9:[],
            rowexport_type_M10:[],
            first_fix_date:"",
            options: ["End Loss Performance", "Width Loss Performance", "Spread Loss Performance"],

            total_compar_ttl_marker:[],
            total_compar_plug12:[],
            total_compar_per_width_loss:[],

            total_compar_end1_2:[],
            total_compar_endless_length_yd:[],
            total_compar_avg_end:[],
            total_compar_spliceloss:[],
            total_compar_splicelossper:[],
            total_compar_totalcutout:[],
            total_compar_totalcutoutper:[],
            total_compar_rement:[],
            total_compar_rement_per:[],
            total_compar_totallenthloss:[],
            total_compar_totallenthlossper:[],
            rowexport_1_h:[],
            rowexport_2_h:[],
            keep_end_plug1_2:[],
            keep_end_plug1_2_table:[],

            open_end_loss_performance:false,
            open_width_loss_performance:false,
            open_spread_loss_performance:false,

            total_sum_endloss_per_week :[],
            total_sum_splice_per_week:[],
            total_sum_cut_per_week:[],
            total_sum_rem_per_week:[],
            total_sum_ttl_marker_per_week:[],

            result_actual_end:[],
            result_actual_splice:[],
            result_actual_cut:[],
            result_actual_rem:[],

            rowexport_type_1_h:[],
            rowexport_type_2_h:[],
            rowexport_type_3_h:[],
            rowexport_type_4_h:[],
            rowexport_type_5_h:[],
            rowexport_type_6_h:[],
            rowexport_type_7_h:[],
            rowexport_type_8_h:[],
            rowexport_type_9_h:[],
            rowexport_type_10_h:[],
            rowexport_type_11_h:[],
            rowexport_type_12_h:[],
            rowexport_type_13_h:[],
            rowexport_type_14_h:[],
            rowexport_type_15_h:[],
            rowexport_type_16_h:[],
            rowexport_type_17_h:[],
            rowexport_type_18_h:[],
            rowexport_type_19_h:[],
            rowexport_type_20_h:[],
            rowexport_type_M1_h:[],
            rowexport_type_M2_h:[],
            rowexport_type_M3_h:[],
            rowexport_type_M4_h:[],
            rowexport_type_M5_h:[],
            rowexport_type_M6_h:[],
            rowexport_type_M7_h:[],
            rowexport_type_M8_h:[],
            rowexport_type_M9_h:[],
            rowexport_type_M10_h:[],

            sum_ttl_marker_h:[],
            row_result_t1:[],
            row_result_t2:[],
            row_result_t3:[],
            row_result_t4:[],
            row_result_t5:[],
            row_result_t6:[],
            row_result_t7:[],
            row_result_t8:[],
            row_result_t9:[],
            row_result_t10:[],
            row_result_t11:[],
            row_result_t12:[],
            row_result_t13:[],
            row_result_t14:[],
            row_result_t15:[],
            row_result_t16:[],
            row_result_t17:[],
            row_result_t18:[],
            row_result_t19:[],
            row_result_t20:[],
            row_result_tM1:[],
            row_result_tM2:[],
            row_result_tM3:[],
            row_result_tM4:[],
            row_result_tM5:[],
            row_result_tM6:[],
            row_result_tM7:[],
            row_result_tM8:[],
            row_result_tM9:[],
            row_result_tM10:[],

            grand_total_ttl_maker:[],
            grand_total_end_loss:[],
            grand_total_splice_loss:[],
            grand_total_cut_out_loss:[],
            grand_total_remnants_loss:[],
            grand_total_length_loss:[],


            row_result_endloss_t1:[],
            row_result_endloss_t2:[],
            row_result_endloss_t3:[],
            row_result_endloss_t4:[],
            row_result_endloss_t5:[],
            row_result_endloss_t6:[],
            row_result_endloss_t7:[],
            row_result_endloss_t8:[],
            row_result_endloss_t9:[],
            row_result_endloss_t10:[],
            row_result_endloss_t11:[],
            row_result_endloss_t12:[],
            row_result_endloss_t13:[],
            row_result_endloss_t14:[],
            row_result_endloss_t15:[],
            row_result_endloss_t16:[],
            row_result_endloss_t17:[],
            row_result_endloss_t18:[],
            row_result_endloss_t19:[],
            row_result_endloss_t20:[],
            row_result_endloss_tM1:[],
            row_result_endloss_tM2:[],
            row_result_endloss_tM3:[],
            row_result_endloss_tM4:[],
            row_result_endloss_tM5:[],
            row_result_endloss_tM6:[],
            row_result_endloss_tM7:[],
            row_result_endloss_tM8:[],
            row_result_endloss_tM9:[],
            row_result_endloss_tM10:[],

            print_status:false,

            



            actual_width_loss:[],
            actual_total_length_loss:[],
            
            
           



            
            column_main:[
                {
                    col_name:"A"
                },
                {
                    col_name:"B"
                },
                {
                    col_name:"C"
                },
                {
                    col_name:"D"
                },
                {
                    col_name:"E"
                },
                {
                    col_name:"F"
                },
                {
                    col_name:"G"
                },
                {
                    col_name:"H"
                },
                {
                    col_name:"I"
                },
                {
                    col_name:"J"
                },
                {
                    col_name:"K"
                },
                {
                    col_name:"L"
                },
                {
                    col_name:"M"
                },
                {
                    col_name:"N"
                },
                {
                    col_name:"O"
                },
                {
                    col_name:"P"
                },
                {
                    col_name:"Q"
                },
                {
                    col_name:"R"
                },
                {
                    col_name:"S"
                },
                {
                    col_name:"T"
                },
                {
                    col_name:"U"
                },
                {
                    col_name:"V"
                },
                {
                    col_name:"W"
                },
                {
                    col_name:"X"
                },
                {
                    col_name:"Y"
                },
                {
                    col_name:"Z"
                },
                {
                    col_name:"AA"
                },
                {
                    col_name:"AB"
                },
                {
                    col_name:"AC"
                },
                {
                    col_name:"AD"
                },
                {
                    col_name:"AE"
                },
                {
                    col_name:"AF"
                },
            ],
            column_second:[
                {
                    col_name:"A"
                },
                {
                    col_name:"B"
                },
                {
                    col_name:"C"
                },
                {
                    col_name:"D"
                },
                {
                    col_name:"E"
                },
                {
                    col_name:"F"
                },
                {
                    col_name:"G"
                },
                {
                    col_name:"H"
                },
                {
                    col_name:"I"
                },
                {
                    col_name:"J"
                },
                {
                    col_name:"K"
                },
                {
                    col_name:"L"
                },
                {
                    col_name:"M"
                },
                {
                    col_name:"N"
                },
                {
                    col_name:"O"
                },
                {
                    col_name:"P"
                },
                {
                    col_name:"Q"
                },
                {
                    col_name:"R"
                },
                {
                    col_name:"S"
                },
                {
                    col_name:"T"
                },
                {
                    col_name:"U"
                },
                {
                    col_name:"V"
                },
                {
                    col_name:"W"
                },
                {
                    col_name:"X"
                },
                {
                    col_name:"Y"
                },
                {
                    col_name:"Z"
                },
                {
                    col_name:"AA"
                },
                {
                    col_name:"AB"
                },
               
            ],
            column_third:[
              {
                col_name:"A"
              },
              {
                col_name:"B"
              },
              {
                col_name:"C"
              },
              {
                col_name:"D"
              },
              {
                col_name:"E"
              },
              {
                col_name:"F"
              },
              {
                col_name:"G"
              },
              {
                col_name:"H"
              },
              {
                col_name:"I"
              },
              {
                col_name:"J"
              },
              {
                col_name:"K"
              },
              {
                col_name:"L"
              },
              {
                col_name:"M"
              },
              {
                col_name:"N"
              },
              {
                col_name:"O"
              },
              {
                col_name:"P"
              },
           
            ],
            forth_column:[
               {
                col_name:"A"
              },
              {
                col_name:"B"
              },
              {
                col_name:"C"
              },
              {
                col_name:"D"
              },
              {
                col_name:"E"
              },
              {
                col_name:"F"
              },
              {
                col_name:"G"
              },
              {
                col_name:"H"
              },
              {
                col_name:"I"
              },
              {
                col_name:"J"
              },
              {
                col_name:"K"
              },
              {
                col_name:"L"
              },
              {
                col_name:"M"
              },
              {
                col_name:"N"
              },
              {
                col_name:"O"
              },
              {
                col_name:"P"
              },
              {
                col_name:"Q"
              },
              {
                col_name:"R"
              },
              {
                col_name:"S"
              },
              {
                col_name:"T"
              },
              {
                col_name:"U"
              },
              {
                col_name:"V"
              },
              {
                col_name:"W"
              },
              {
                col_name:"X"
              },
              {
                col_name:"Y"
              },
              {
                col_name:"Z"
              },
              {
                col_name:"AA"
              },
              {
                col_name:"AB"
              },
              {
                col_name:"AC"
              },
              {
                col_name:"AD"
              },
              {
                col_name:"AE"
              },
              {
                col_name:"AF"
              },
              {
                col_name:"AG"
              },
              {
                col_name:"AH"
              },
              {
                col_name:"AI"
              },
              {
                col_name:"AJ"
              },
              {
                col_name:"AK"
              },
              {
                col_name:"AL"
              },
              {
                col_name:"AM"
              },
              {
                col_name:"AN"
              },
              {
                col_name:"AO"
              },
              {
                col_name:"AP"
              },
              {
                col_name:"AQ"
              },
              {
                col_name:"AR"
              },
              {
                col_name:"AS"
              },
              {
                col_name:"AT"
              },
              {
                col_name:"AU"
              },
              {
                col_name:"AV"
              },
              {
                col_name:"AW"
              },
              {
                col_name:"AX"
              },
              {
                col_name:"AY"
              },
              {
                col_name:"AZ"
              },
              {
                col_name:"BA"
              },
              {
                col_name:"BB"
              },
              {
                col_name:"BC"
              },
              {
                col_name:"BD"
              },
              {
                col_name:"BE"
              },
              {
                col_name:"BF"
              },
              {
                col_name:"BG"
              },
              {
                col_name:"BH"
              },
              {
                col_name:"BI"
              },
              {
                col_name:"BJ"
              },
              {
                col_name:"BK"
              },
              {
                col_name:"BL"
              },
              {
                col_name:"BM"
              },
              {
                col_name:"BN"
              },
              {
                col_name:"BO"
              },
              {
                col_name:"BP"
              },
              {
                col_name:"BQ"
              },
              {
                col_name:"BR"
              },
              {
                col_name:"BS"
              },
              {
                col_name:"BT"
              },
              {
                col_name:"BU"
              },
              {
                col_name:"BV"
              },
              {
                col_name:"BW"
              },
              {
                col_name:"BX"
              },
          
            ]
            
        }
    },
    mounted() {
    let org = this.$q.localStorage.getItem("org");
    this.org = org
    axios.get(this.$api_url + "/mainso.php/get_date2")
      .then((resp)=>{
        this.first_fix_date = (resp.data.data[0].START_DATE)
      }).catch((error)=>{
        console.log(error)
      })
    
   
    
    },
    methods:{
            
        async exportexcel(){
          this.$q.loading.show({
          spinner: QSpinnerGears,
          spinnerColor: 'wthite',
          spinnerSize: 180,
          backgroundColor: 'black',
         
        })
            this.print_status = false
            this.date_use_data  = []
            this.date_use_data2  = []
            this.rowexport_type_1 = []
            this.rowexport_type_2=[]
            this.rowexport_type_3=[]
            this.rowexport_type_4=[]
            this.rowexport_type_5=[]
            this.rowexport_type_6=[]
            this.rowexport_type_7=[]
            this.rowexport_type_8=[]
            this.rowexport_type_9=[]
            this.rowexport_type_10=[]
            this.rowexport_type_11=[]
            this.rowexport_type_12=[]
            this.rowexport_type_13=[]
            this.rowexport_type_14=[]
            this.rowexport_type_15=[]
            this.rowexport_type_16=[]
            this.rowexport_type_17=[]
            this.rowexport_type_18=[]
            this.rowexport_type_19=[]
            this.rowexport_type_20=[]
            this.rowexport_type_M1=[]
            this.rowexport_type_M2=[]
            this.rowexport_type_M3=[]
            this.rowexport_type_M4=[]
            this.rowexport_type_M5=[]
            this.rowexport_type_M6=[]
            this.rowexport_type_M7=[]
            this.rowexport_type_M8=[]
            this.rowexport_type_M9=[]
            this.rowexport_type_M10=[]
            this.total_compar_ttl_marker=[]
            this.total_compar_plug12=[]
            this.total_compar_per_width_loss=[]
            this.total_compar_end1_2=[]
            this.total_compar_endless_length_yd=[]
            this.total_compar_avg_end=[]
            this.total_compar_spliceloss=[]
            this.total_compar_splicelossper=[]
            this.total_compar_totalcutout=[]
            this.total_compar_totalcutoutper=[]
            this.total_compar_rement=[]
            this.total_compar_rement_per=[]
            this.total_compar_totallenthloss=[]
            this.total_compar_totallenthlossper=[]
            this.rowexport_1_h=[]
            this.rowexport_2_h=[]

       

            this.total_sum_endloss_per_week = [],
            this.total_sum_splice_per_week = [],
            this.total_sum_cut_per_week = [],
            this.total_sum_rem_per_week = [],
            this.total_sum_ttl_marker_per_week =[],

            this.result_actual_end=[],
            this.result_actual_splice=[],
            this.result_actual_cut=[],
            this.result_actual_rem=[],

            this.row_result_t1=[]
            this.row_result_t2=[]
            this.row_result_t3=[]
            this.row_result_t4=[]
            this.row_result_t5=[]
            this.row_result_t6=[]
            this.row_result_t7=[]
            this.row_result_t8=[]
            this.row_result_t9=[]
            this.row_result_t10=[]
            this.row_result_t11=[]
            this.row_result_t12=[]
            this.row_result_t13=[]
            this.row_result_t14=[]
            this.row_result_t15=[]
            this.row_result_t16=[]
            this.row_result_t17=[]
            this.row_result_t18=[]
            this.row_result_t19=[]
            this.row_result_t20=[]
            this.row_result_tM1=[]
            this.row_result_tM2=[]
            this.row_result_tM3=[]
            this.row_result_tM4=[]
            this.row_result_tM5=[]
            this.row_result_tM6=[]
            this.row_result_tM7=[]
            this.row_result_tM8=[]
            this.row_result_tM9=[]
            this.row_result_tM10=[]

            this.keep_end_plug1_2 = []
            this.keep_end_plug1_2_table = []

            this.row_result_endloss_t1=[]
            this.row_result_endloss_t2=[]
            this.row_result_endloss_t3=[]
            this.row_result_endloss_t4=[]
            this.row_result_endloss_t5=[]
            this.row_result_endloss_t6=[]
            this.row_result_endloss_t7=[]
            this.row_result_endloss_t8=[]
            this.row_result_endloss_t9=[]
            this.row_result_endloss_t10=[]
            this.row_result_endloss_t11=[]
            this.row_result_endloss_t12=[]
            this.row_result_endloss_t13=[]
            this.row_result_endloss_t14=[]
            this.row_result_endloss_t15=[]
            this.row_result_endloss_t16=[]
            this.row_result_endloss_t17=[]
            this.row_result_endloss_t18=[]
            this.row_result_endloss_t19=[]
            this.row_result_endloss_t20=[]
            this.row_result_endloss_tM1=[]
            this.row_result_endloss_tM2=[]
            this.row_result_endloss_tM3=[]
            this.row_result_endloss_tM4=[]
            this.row_result_endloss_tM5=[]
            this.row_result_endloss_tM6=[]
            this.row_result_endloss_tM7=[]
            this.row_result_endloss_tM8=[]
            this.row_result_endloss_tM9=[]
            this.row_result_endloss_tM10=[]

            this.actual_width_loss = []
            this.actual_total_length_loss=[]
/* 
     var startDateX = moment(this.start_h,'YYYY-MM-DD').add(3, 'days')
     var endDateX = moment(this.end_h,'YYYY-MM-DD').add(3, 'days')
  
     var weekDayName =  moment(startDateX).format('dddd');
     var weekDayNameX =  moment(endDateX).format('dddd');
            
       
          console.log(weekDayName);
          console.log(weekDayNameX)

          if(weekDayName == 'Sunday' || weekDayNameX == 'Sunday'){
             this.$q.notify({
            message: "Don't Choose Sunday.",

            color: "red",
          });
          }
 */
        
            
      const workbook = new Excel.Workbook();
      workbook.creator = "Nanyang";
      workbook.lastModifiedBy = "Admin";
      workbook.created = new Date(2021, 8, 30);
      workbook.modified = new Date();
      workbook.lastPrinted = new Date(2021, 7, 27);
     
      const ws = workbook.addWorksheet(this.org);
      const ws1 = workbook.addWorksheet("Total Compare");
      const ws2 = workbook.addWorksheet("Spread Loss By Floor-h");
      const ws3 = workbook.addWorksheet("End Loss Performance");
      const ws4 = workbook.addWorksheet("Width Loss Performance");
      const ws5 = workbook.addWorksheet("Spread Loss Performance");

   
 


   

    const params  = new FormData()
    params.append("start",this.start_date)
    params.append("end",this.end_date)
    let org = this.$q.localStorage.getItem("org");
    params.append("org",org)


    await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data_weekly_analy',
        data:params,
      }).then((resp)=>{
         console.log(resp.data)
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
                e.ISSUE - (e.TTL_MARKER_YD * e.WEIGHT_YD) / 100
              ).toFixed(2),
            })
          
          });
          }
      })
  
   
    //sheet2
    ws.columns = [
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
        { key: "S", width: 6.7 },
        { key: "T", width: 13 },
        { key: "U", width: 6.7 },
        { key: "V", width: 6.7 },
        { key: "W", width: 13 },
        { key: "X", width: 6.7 },
        { key: "AC", width: 15.13 },
      ];

        //control grid
    for(var ax = 0; ax<this.column_second.length; ax++){
         for(var bz = 1; bz<400;  bz++){
        ws.getCell(this.column_second[ax].col_name+bz).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
      
       ws.getCell(this.column_second[ax].col_name+bz).alignment = {
          horizontal: "center",
          vertical: "middle",
        };

        ws.getCell(this.column_second[ax].col_name+bz).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 

      }
    }
       ws.getCell("A1").value = "Spread Loss Summary Analysis :" + this.org
       ws.getCell("A1").alignment = {
          horizontal:"left"
       };
       ws.getCell("A1").font = {
        name: 'Angsana New',
        color: { argb: 'FF3163CF' },
        family: 4,
        size: 24,
        bold: false
      };
       ws.mergeCells("A1:AB1")

 

       ws.getCell("A2").value = "Period"
       ws.mergeCells("A2:C2")

       ws.getCell("D2").value = this.start + " - " + this.end
       ws.mergeCells("D2:E2")

       ws.getCell("F2").value = "FARBRIC TYPE"
       ws.mergeCells("F2:G2")

       ws.getCell("H2").value = "MARKER"
       ws.mergeCells("H2:J2")

       ws.getCell("K2").value = "Width Loss"
       ws.mergeCells("K2:M2")

       ws.getCell("N2").value = "End Loss"
       ws.mergeCells("N2:Q2")

       ws.getCell("R2").value = "Splice Loss"
       ws.mergeCells("R2:T2")

       ws.getCell("U2").value = "Cut Out Loss"
       ws.mergeCells("U2:W2")

       ws.getCell("X2").value = "Remnants Loss"
       ws.mergeCells("X2:Z2")

       ws.getCell("AA2").value = "Total Length Loss"
       ws.mergeCells("AA2:AB2")

       ws.getCell("A3").value = "Date"
       ws.mergeCells("A3:A4")

       ws.getCell("B3").value = "Table"
       ws.mergeCells("B3:B4")

       ws.getCell("C3").value = "S / O no."
       ws.mergeCells("C3:C4")

       ws.getCell("D3").value = "Customer"
       ws.mergeCells("D3:D4")

       ws.getCell("E3").value = "GM08 no."
       ws.mergeCells("E3:E4")

       ws.getCell("F3").value = "Fabric Type"
       ws.mergeCells("F3:F4")

        ws.getCell("G3").value = "Observe";
        ws.getCell("G4").value = "Width";

        ws.getCell("H3").value = "Width";
        ws.getCell("H4").value = "(inch)";

        ws.getCell("I3").value = "Length";
        ws.getCell("I4").value = "(yd)";

        ws.getCell("J3").value = "Ttl Marker";
        ws.getCell("J4").value = "Length(yd)";

        ws.getCell("K3").value = "Plug 1+2";
        ws.getCell("K4").value = "(inch)";

        ws.getCell("L3").value = 1.50/100;
        ws.getCell("L3").numFmt = "0.00%";   
        ws.getCell("L4").value = "WIdth Loss";

        ws.getCell("M3").value = "*Remark";
        ws.getCell("M4").value = "Code";

        ws.getCell("N3").value = "Plug1+2";
        ws.getCell("N4").value = "(inch)";

        ws.getCell("O3").value = "Length";
        ws.getCell("O4").value = "(yd)";

        ws.getCell("P3").value = "%";
        ws.getCell("P4").value = "End";

        ws.getCell("Q3").value = "*Remark";
        ws.getCell("Q4").value = "Code";

        ws.getCell("R3").value = "Length";
        ws.getCell("R4").value = "(yd)";

        ws.getCell("S3").value = "%";
        ws.getCell("S4").value = "Splice";

        ws.getCell("T3").value = "*Remark";
        ws.getCell("T4").value = "code";

        ws.getCell("U3").value = "Length";
        ws.getCell("U4").value = "(yd)";

        ws.getCell("V3").value = "%";
        ws.getCell("V4").value = "Cut Out";

        ws.getCell("W3").value = "*Remark";
        ws.getCell("W4").value = "code";

        ws.getCell("X3").value = "Length";
        ws.getCell("X4").value = "(yd)";

        ws.getCell("Y3").value = "%";
        ws.getCell("Y4").value = "Remnants";

        ws.getCell("Z3").value = "*Remark";
        ws.getCell("Z4").value = "code";

        ws.getCell("AA3").value = "Length";
        ws.getCell("AA4").value = "(yd)";

        ws.getCell("AB3").value = "% of";
        ws.getCell("AB4").value = "Total Length";

  
      this.result_sum_ttl_width = 0
      this.result_sum_ttl_width_per = 0
      this.result_sum_ttl_plug12_end = 0
      for(var ax = 0; ax<this.rowexport.length; ax++){
         this.sum_ttl_width1_2 = 0
          this.sum_ttl_width1_2 = this.rowexport[ax].ttl_marker * this.rowexport[ax].plug12
          this.result_sum_ttl_width = Number(this.result_sum_ttl_width) + Number(this.sum_ttl_width1_2)

          this.sum_ttl_width_per = 0
          this.sum_ttl_width_per = this.rowexport[ax].ttl_marker * this.rowexport[ax].widthloss
          this.result_sum_ttl_width_per = Number(this.result_sum_ttl_width_per) + this.sum_ttl_width_per

         /*  this.sum_ttl_plug12_end = 0
          this.sum_ttl_plug12_end = this.rowexport[ax].ttl_marker * this.rowexport[ax]. */


      }


      for(var ax = 0; ax<this.rowexport.length; ax++){
        if(this.rowexport[ax].gm_table == "A1"){
          this.rowexport_type_1.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
        }else if(this.rowexport[ax].gm_table == "A2"){
        this.rowexport_type_2.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
        }else if(this.rowexport[ax].gm_table == "A3"){
        this.rowexport_type_3.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
       }else if(this.rowexport[ax].gm_table == "A4"){
        this.rowexport_type_4.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "A5"){
        this.rowexport_type_5.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "A6"){
        this.rowexport_type_6.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "A7"){
        this.rowexport_type_7.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "A8"){
        this.rowexport_type_8.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "A9"){
        this.rowexport_type_9.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "A10"){
        this.rowexport_type_10.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "A11"){
        this.rowexport_type_11.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "A12"){
        this.rowexport_type_12.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "A13"){
        this.rowexport_type_13.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "A14"){
        this.rowexport_type_14.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "A15"){
        this.rowexport_type_15.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "A16"){
        this.rowexport_type_16.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "A17"){
        this.rowexport_type_17.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "A18"){
        this.rowexport_type_18.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "A19"){
        this.rowexport_type_19.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "A20"){
        this.rowexport_type_20.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "M1"){
        this.rowexport_type_M1.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "M2"){
        this.rowexport_type_M2.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "M3"){
        this.rowexport_type_M3.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "M4"){
        this.rowexport_type_M4.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "M5"){
        this.rowexport_type_M5.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "M6"){
        this.rowexport_type_M6.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "M7"){
        this.rowexport_type_M7.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "M8"){
        this.rowexport_type_M8.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "M9"){
        this.rowexport_type_M9.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }else if(this.rowexport[ax].gm_table == "M10"){
        this.rowexport_type_M10.push({
              mu_date: this.rowexport[ax].mu_date,
              gm_table: this.rowexport[ax].gm_table,
              so_no_doc: this.rowexport[ax].so_no_doc,
              customer: this.rowexport[ax].customer,
              gm_no: this.rowexport[ax].gm_no,
              gm_no_short: this.rowexport[ax].gm_no_short,
              fabric_type: this.rowexport[ax].fabric_type,
              obs_width:this.rowexport[ax].obs_width,
              width_inch:this.rowexport[ax].width_inch,
              length_ydb:this.rowexport[ax].length_ydb,
              ttl_marker:this.rowexport[ax].ttl_marker,
              plug12:this.rowexport[ax].plug12,
              widthloss:this.rowexport[ax].widthloss,
              endless_length_yd:this.rowexport[ax].endless_length_yd,
              avg_end:this.rowexport[ax].avg_end,
              end1_2:this.rowexport[ax].end1_2,
              spliceloss:this.rowexport[ax].spliceloss,
              splicelossper:this.rowexport[ax].splicelossper,
              totalcutout:this.rowexport[ax].totalcutout,
              totalcutoutper:this.rowexport[ax].totalcutoutper,
              rement:this.rowexport[ax].rement,
              rement_per:this.rowexport[ax].rement_per,
              totallenthloss:this.rowexport[ax].totallenthloss,
              totallenthlossper:this.rowexport[ax].totallenthlossper,
              widthloss_code:this.rowexport[ax].widthloss_code,
              avg_end_code:this.rowexport[ax].avg_end_code,
              per_splice_loss_code:this.rowexport[ax].per_splice_loss_code,
              total_cut_out_per_code:this.rowexport[ax].total_cut_out_per_code,
              remnants_loss_code:this.rowexport[ax].remnants_loss_code,

          })
      }
      }
      
       var count = 5
       var count_row_grand = 0
        var countx = 5 + Number(this.rowexport_type_1.length)
      
      this.sum_ttl_marker_t1 = 0
      this.sum_plug12_t1 = 0
      this.sum_per_width_loss_t1 = 0
      this.sum_end1_2_t1 = 0
      this.sum_endless_length_yd_t1= 0
      this.sum_avg_end_t1 = 0
      this.sum_spliceloss_t1 = 0
      this.sum_splicelossper_t1 = 0
      this.sum_totalcutout_t1 = 0
      this.sum_totalcutoutper_t1 = 0
      this.sum_rement_t1 = 0
      this.sum_rement_per_t1 = 0
      this.sum_totallenthloss_t1 = 0 
      this.sum_totallenthlossper_t1 = 0
      this.summary_widthloss_t1 = 0
      this.summary_plug12_t1 = 0
      this.summary_end1_2_t1 = 0
      this.sum_pre_plug1_2_end = 0
      this.sum_plug1_2_end = 0
      
      
    
     this.sum_all_plug1_2_end_total_t1 = 0
   
      
     if(this.rowexport_type_1.length < 1){
      count_row_grand++
     }else{
      
      for (var ax = 0; ax<this.rowexport_type_1.length; ax++){ //length +20
     
      ws.getCell("A"+count).value = this.rowexport_type_1[ax].mu_date

      ws.getCell("B"+count).value = this.rowexport_type_1[ax].gm_table

      ws.getCell("C"+count).value = Number(this.rowexport_type_1[ax].gm_no_short)
      ws.getCell("C"+count).numFmt = '0.00';

      ws.getCell("D"+count).value = this.rowexport_type_1[ax].customer

      ws.getCell("E"+count).value = this.rowexport_type_1[ax].gm_no

      ws.getCell("F"+count).value = this.rowexport_type_1[ax].fabric_type

      ws.getCell("G"+count).value = Number(this.rowexport_type_1[ax].obs_width)
      ws.getCell("G"+count).numFmt = '0.00';

      ws.getCell("H"+count).value = Number(this.rowexport_type_1[ax].width_inch)
      ws.getCell("H"+count).numFmt = '0.00';


      ws.getCell("I"+count).value = Number(this.rowexport_type_1[ax].length_ydb)
      ws.getCell("I"+count).numFmt = '0.00';

      ws.getCell("J"+count).value = Number(this.rowexport_type_1[ax].ttl_marker)
      ws.getCell("J"+count).numFmt = '0.00';

    
      ws.getCell("K"+count).value = Number(this.rowexport_type_1[ax].plug12)
      ws.getCell("K"+count).numFmt = '0.00';

      ws.getCell("L"+count).value = this.rowexport_type_1[ax].widthloss/100
      ws.getCell("L"+count).numFmt = '0.00%';

      ws.getCell("M"+count).value = this.rowexport_type_1[ax].widthloss_code
    

      ws.getCell("O"+count).value = Number(this.rowexport_type_1[ax].endless_length_yd)
      ws.getCell("O"+count).numFmt = '0.00';

      ws.getCell("P"+count).value = this.rowexport_type_1[ax].avg_end/100
      ws.getCell("P"+count).numFmt = '0.00%';

      ws.getCell("Q"+count).value = this.rowexport_type_1[ax].avg_end_code
      

      ws.getCell("R"+count).value = Number(this.rowexport_type_1[ax].spliceloss)
      ws.getCell("R"+count).numFmt = '0.00';

      ws.getCell("S"+count).value = this.rowexport_type_1[ax].splicelossper/100
      ws.getCell("S"+count).numFmt = '0.00%';

      ws.getCell("T"+count).value = this.rowexport_type_1[ax].per_splice_loss_code
   

      ws.getCell("U"+count).value = Number(this.rowexport_type_1[ax].totalcutout)
      ws.getCell("U"+count).numFmt = '0.00';

      ws.getCell("V"+count).value = this.rowexport_type_1[ax].totalcutoutper/100
      ws.getCell("V"+count).numFmt = '0.00%';

      ws.getCell("W"+count).value = this.rowexport_type_1[ax].total_cut_out_per_code

      ws.getCell("X"+count).value = Number(this.rowexport_type_1[ax].rement)
      ws.getCell("X"+count).numFmt = '0.00';

      ws.getCell("Y"+count).value = this.rowexport_type_1[ax].rement_per/100
      ws.getCell("Y"+count).numFmt = '0.00%';

      ws.getCell("Z"+count).value = this.rowexport_type_1[ax].remnants_loss_code

      ws.getCell("AA"+count).value = Number(this.rowexport_type_1[ax].totallenthloss)
      ws.getCell("AA"+count).numFmt = '0.00';

      ws.getCell("AB"+count).value = this.rowexport_type_1[ax].totallenthlossper/100
      ws.getCell("AB"+count).numFmt = '0.00%';
      
      this.sum_plug12_t1 = 0
      
      this.sum_ttl_marker_t1 = Number(this.sum_ttl_marker_t1) + Number(this.rowexport_type_1[ax].ttl_marker)
      this.sum_plug12_t1 = this.rowexport_type_1[ax].plug12 * this.rowexport_type_1[ax].ttl_marker
      this.summary_plug12_t1 = Number(this.summary_plug12_t1) + Number(this.sum_plug12_t1)
      this.sum_per_width_loss_t1 = this.rowexport_type_1[ax].widthloss * this.rowexport_type_1[ax].ttl_marker
      this.summary_widthloss_t1 = Number(this.summary_widthloss_t1) + Number(this.sum_per_width_loss_t1)
      this.sum_end1_2_t1 = this.rowexport_type_1[ax].end1_2 * this.rowexport_type_1[ax].ttl_marker
      this.summary_end1_2_t1 = Number(this.summary_end1_2_t1) + Number(this.sum_end1_2_t1)
      this.sum_endless_length_yd_t1 = Number(this.sum_endless_length_yd_t1) + Number(this.rowexport_type_1[ax].endless_length_yd)
      this.sum_spliceloss_t1 = Number(this.sum_spliceloss_t1) + Number(this.rowexport_type_1[ax].spliceloss)
      this.sum_totalcutout_t1 = Number(this.sum_totalcutout_t1) + Number(this.rowexport_type_1[ax].totalcutout)
      this.sum_rement_t1 = Number(this.sum_rement_t1) + Number(this.rowexport_type_1[ax].rement)
      this.sum_totallenthloss_t1_first = 0
      this.sum_totallenthloss_t1_first = Number(this.rowexport_type_1[ax].endless_length_yd) + Number(this.rowexport_type_1[ax].spliceloss)
      +Number(this.rowexport_type_1[ax].totalcutout) + Number(this.rowexport_type_1[ax].rement)
      this.sum_totallenthloss_t1 = Number(this.sum_totallenthloss_t1) + Number(this.sum_totallenthloss_t1_first)

       this.sum_totallenthloss_per_t1_first = 0
      this.sum_totallenthloss_per_t1_first = Number(this.rowexport_type_1[ax].avg_end) + Number(this.rowexport_type_1[ax].splicelossper)
      +Number(this.rowexport_type_1[ax].totalcutoutper) + Number(this.rowexport_type_1[ax].rement_per)
      this.sum_totallenthloss_per_t1 = Number(this.sum_totallenthloss_t1) + Number(this.sum_totallenthloss_per_t1_first)
      this.sum_number_1_end_t1 = 0 
      this.sum_number_1_end_t1 = this.rowexport_type_1[ax].endless_length_yd * 36
      this.sum_number_2_end_t1 = 0
      this.sum_number_2_end_t1 = this.rowexport_type_1[ax].ttl_marker/this.rowexport_type_1[ax].length_ydb
      this.sum_plug1_2_end_t1 = 0
      this.sum_plug1_2_end_t1 = this.sum_number_1_end_t1/this.sum_number_2_end_t1


      this.sum_plug1_2_end_total_t1 = 0
     
      this.sum_plug1_2_end_total_t1 = this.sum_plug1_2_end_t1 * this.rowexport_type_1[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t1 = Number(this.sum_all_plug1_2_end_total_t1) + Number(this.sum_plug1_2_end_total_t1)

      this.grand_total_ttl_maker.push({
        value:this.sum_ttl_marker_t1
      })

      this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t1
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t1
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t1
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t1
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t1
      })
     
     this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t1,
        value_table:1
      })

      ws.getCell("N"+count).value = Number(this.sum_plug1_2_end_t1)
      ws.getCell("N"+count).numFmt = '0.00';

      count++
         }
       
         if(this.sum_ttl_marker_t1 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t1
         })
         }else{  
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
 
 this.sum_plug12_t1 = this.summary_plug12_t1 / this.sum_ttl_marker_t1 
        if(this.sum_plug12_t1 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t1
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t1 =  this.summary_widthloss_t1 / this.sum_ttl_marker_t1 
        if(this.sum_per_width_loss_t1 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t1
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t1 =  this.summary_end1_2_t1 / this.sum_ttl_marker_t1 //M
      if(this.sum_end1_2_t1 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t1
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t1 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t1
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t1 = this.sum_endless_length_yd_t1 / this.sum_ttl_marker_t1 *100 //P
      if(this.sum_avg_end_t1 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t1
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t1 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t1
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t1 = this.sum_spliceloss_t1 / this.sum_ttl_marker_t1 *100
      if(this.sum_splicelossper_t1 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t1
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t1 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t1
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t1 = this.sum_totalcutout_t1 / this.sum_ttl_marker_t1 *100
      if(this.sum_totalcutoutper_t1 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t1
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t1 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t1
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t1 = this.sum_rement_t1 / this.sum_ttl_marker_t1 *100
  if(this.sum_rement_per_t1 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t1
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t1
    if(this.sum_totallenthloss_t1 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t1
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert1 = this.sum_totallenthlosst1 / this.sum_ttl_marker_t1
    if(this.sum_totallenthloss_per_t1 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthloss_per_t1
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

   
      ws.getCell("A"+countx).value = "Total Table A1"  
      ws.mergeCells("A"+countx+":"+"I"+countx)
     
      if(this.sum_ttl_marker_t1 > 0 && isNaN(this.sum_ttl_marker_t1)==false && isFinite(this.sum_ttl_marker_t1)==true){
      ws.getCell("J"+countx).value = Number(this.sum_ttl_marker_t1)
      ws.getCell("J"+count).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx).value = 0
      ws.getCell("J"+count).numFmt = '0.00';
      }

      if(this.sum_plug12_t1 > 0 && isNaN(this.sum_plug12_t1)==false && isFinite(this.sum_plug12_t1)==true){
      ws.getCell("K"+countx).value = Number(this.sum_plug12_t1)
      ws.getCell("K"+countx).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx).value = 0
      ws.getCell("K"+countx).numFmt = '0.00';
      }

     
      if(this.sum_per_width_loss_t1 > 0 && isNaN(this.sum_per_width_loss_t1)==false && isFinite(this.sum_per_width_loss_t1)==true){
      
      ws.getCell("L"+countx).value = this.sum_per_width_loss_t1/100
      ws.getCell("L"+countx).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx).value = 0.00/100
      ws.getCell("L"+countx).numFmt = '0.00%'; 
      }

      
      this.total_result_plug1_2_t1 = 0
      console.log(this.sum_all_plug1_2_end_total_t1)
      console.log(this.sum_ttl_marker_t1)
      this.total_result_plug1_2_t1 = this.sum_all_plug1_2_end_total_t1 / this.sum_ttl_marker_t1

      this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t1
      })
      if(this.total_result_plug1_2_t1 > 0 && isNaN(this.total_result_plug1_2_t1)==false && isFinite(this.total_result_plug1_2_t1)==true){
      ws.getCell("N"+countx).value = Number(this.total_result_plug1_2_t1)
      ws.getCell("N"+countx).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx).value = 0.00
      ws.getCell("N"+countx).numFmt = '0.00';
      }
      
      

      if(this.sum_endless_length_yd_t1 > 0 && isNaN(this.sum_endless_length_yd_t1)==false && isFinite(this.sum_endless_length_yd_t1)==true){
      ws.getCell("O"+countx).value = Number(this.sum_endless_length_yd_t1)
      ws.getCell("O"+countx).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx).value = 0.00
      ws.getCell("O"+countx).numFmt = '0.00';
      }

      if(this.sum_avg_end_t1 > 0 && isNaN(this.sum_avg_end_t1)==false && isFinite(this.sum_avg_end_t1)==true){
      ws.getCell("P"+countx).value = this.sum_avg_end_t1/100
      ws.getCell("P"+countx).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx).value = 0/100
      ws.getCell("P"+countx).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t1 > 0 && isNaN(this.sum_spliceloss_t1)==false && isFinite(this.sum_spliceloss_t1)==true){
      ws.getCell("R"+countx).value = Number(this.sum_spliceloss_t1)
      ws.getCell("R"+countx).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx).value = 0.00
      ws.getCell("R"+countx).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t1 > 0 && isNaN(this.sum_splicelossper_t1)==false && isFinite(this.sum_splicelossper_t1)==true){
      ws.getCell("S"+countx).value = this.sum_splicelossper_t1/100
      ws.getCell("S"+countx).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx).value = 0/100
      ws.getCell("S"+countx).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t1 > 0 && isNaN(this.sum_totalcutout_t1)==false && isFinite(this.sum_totalcutout_t1)==true){
      ws.getCell("U"+countx).value = Number(this.sum_totalcutout_t1)
      ws.getCell("U"+countx).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx).value = 0.00
      ws.getCell("U"+countx).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t1 > 0 && isNaN(this.sum_totalcutoutper_t1)==false && isFinite(this.sum_totalcutoutper_t1)==true){
      ws.getCell("V"+countx).value = this.sum_totalcutoutper_t1/100
      ws.getCell("V"+countx).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx).value = 0.00/100
      ws.getCell("V"+countx).numFmt = '0.00%';
      }

       if(this.sum_rement_t1 > 0 && isNaN(this.sum_rement_t1)==false && isFinite(this.sum_rement_t1)==true){
      ws.getCell("X"+countx).value = Number(this.sum_rement_t1)
      ws.getCell("X"+countx).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx).value = 0.00
      ws.getCell("X"+countx).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t1 > 0 && isNaN(this.sum_rement_per_t1)==false && isFinite(this.sum_rement_per_t1)==true){
      ws.getCell("Y"+countx).value =  this.sum_rement_per_t1/100
      ws.getCell("Y"+countx).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx).value =  0.00/100
      ws.getCell("Y"+countx).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t1 > 0 && isNaN(this.sum_totallenthloss_t1)==false && isFinite(this.sum_totallenthloss_t1)==true){
      ws.getCell("AA"+countx).value = Number(this.sum_totallenthloss_t1)
      ws.getCell("AA"+countx).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx).value = 0
      ws.getCell("AA"+countx).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t1 = this.sum_totallenthloss_t1 / this.sum_ttl_marker_t1 *100
      console.log(this.last_sum_totallenthloss_per_t1)
      if(this.last_sum_totallenthloss_per_t1 > 0 && isNaN(this.last_sum_totallenthloss_per_t1)==false && isFinite(this.last_sum_totallenthloss_per_t1)==true){
      ws.getCell("AB"+countx).value = this.last_sum_totallenthloss_per_t1/100
      ws.getCell("AB"+countx).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx).value = 0/100
      ws.getCell("AB"+countx).numFmt = '0.00%'; 
      }

    for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
    }

   }

      //A2
       if(this.rowexport_type_1.length == 0){
       
         var count2 = countx

   
       }else{
       var count2 = countx+1
       }
   
        var countx2 =  count2 + 1 + Number(this.rowexport_type_2.length)
      
      this.sum_ttl_marker_t2 = 0
      this.sum_plug12_t2 = 0
      this.sum_per_width_loss_t2 = 0
      this.sum_end1_2_t2 = 0
      this.sum_endless_length_yd_t2= 0
      this.sum_avg_end_t2 = 0
      this.sum_spliceloss_t2 = 0
      this.sum_splicelossper_t2 = 0
      this.sum_totalcutout_t2 = 0
      this.sum_totalcutoutper_t2 = 0
      this.sum_rement_t2 = 0
      this.sum_rement_per_t2 = 0
      this.sum_totallenthloss_t2 =0 
      this.sum_totallenthlossper_t2 = 0
      this.summary_widthloss_t2 = 0
      this.summary_plug12_t2 = 0
      this.summary_end1_2_t2 = 0
   
      
      this.sum_all_plug1_2_end_total_t2 = 0
      if(this.rowexport_type_2.length < 1){
        count_row_grand++
      }else{
          for (var ax = 0; ax<this.rowexport_type_2.length; ax++){ //length +20
      this.sum_number_1_end_t1 = 0 
      this.sum_number_1_end_t1 = this.sum_endless_length_yd_t1 * 36
      this.sum_number_2_end_t1 = 0
      this.sum_number_2_end_t1 = this.sum_ttl_marker_t1/this.rowexport_type_2[ax].length_ydb
  
       this.sum_plug1_2_end_t1 = 0
       this.sum_plug1_2_end_t1 = this.sum_number_1_end_t1/this.sum_number_2_end_t1
      ws.getCell("A"+count2).value = this.rowexport_type_2[ax].mu_date

      ws.getCell("B"+count2).value = this.rowexport_type_2[ax].gm_table

      ws.getCell("C"+count2).value = Number(this.rowexport_type_2[ax].gm_no_short)
      ws.getCell("C"+count2).numFmt = '0.00';

      ws.getCell("D"+count2).value = this.rowexport_type_2[ax].customer

      ws.getCell("E"+count2).value = this.rowexport_type_2[ax].gm_no

      ws.getCell("F"+count2).value = this.rowexport_type_2[ax].fabric_type

      ws.getCell("G"+count2).value = Number(this.rowexport_type_2[ax].obs_width)
      ws.getCell("G"+count2).numFmt = '0.00';

      ws.getCell("H"+count2).value = Number(this.rowexport_type_2[ax].width_inch)
      ws.getCell("H"+count2).numFmt = '0.00';


      ws.getCell("I"+count2).value = Number(this.rowexport_type_2[ax].length_ydb)
      ws.getCell("I"+count2).numFmt = '0.00';

      ws.getCell("J"+count2).value = Number(this.rowexport_type_2[ax].ttl_marker)
      ws.getCell("J"+count2).numFmt = '0.00';

      ws.getCell("K"+count2).value = Number(this.rowexport_type_2[ax].plug12)
      ws.getCell("K"+count2).numFmt = '0.00';

      console.log(this.rowexport_type_2[ax].widthloss)
      ws.getCell("L"+count2).value = this.rowexport_type_2[ax].widthloss/100
      ws.getCell("L"+count2).numFmt = '0.00%';

      ws.getCell("M"+count2).value = this.rowexport_type_2[ax].widthloss_code
    

      


      ws.getCell("O"+count2).value = Number(this.rowexport_type_2[ax].endless_length_yd)
      ws.getCell("O"+count2).numFmt = '0.00';

      ws.getCell("P"+count2).value = this.rowexport_type_2[ax].avg_end/100
      ws.getCell("P"+count2).numFmt = '0.00%';

      ws.getCell("Q"+count2).value = this.rowexport_type_2[ax].avg_end_code
      

      ws.getCell("R"+count2).value = Number(this.rowexport_type_2[ax].spliceloss)
      ws.getCell("R"+count2).numFmt = '0.00';

      ws.getCell("S"+count2).value = this.rowexport_type_2[ax].splicelossper/100
      ws.getCell("S"+count2).numFmt = '0.00%';

      ws.getCell("T"+count2).value = this.rowexport_type_2[ax].per_splice_loss_code
   

      ws.getCell("U"+count2).value = Number(this.rowexport_type_2[ax].totalcutout)
      ws.getCell("U"+count2).numFmt = '0.00';

      ws.getCell("V"+count2).value = this.rowexport_type_2[ax].totalcutoutper/100
      ws.getCell("V"+count2).numFmt = '0.00%';

      ws.getCell("W"+count2).value = this.rowexport_type_2[ax].total_cut_out_per_code

      ws.getCell("X"+count2).value = Number(this.rowexport_type_2[ax].rement)
      ws.getCell("X"+count2).numFmt = '0.00';

      ws.getCell("Y"+count2).value = this.rowexport_type_2[ax].rement_per/100
      ws.getCell("Y"+count2).numFmt = '0.00%';

      ws.getCell("Z"+count2).value = this.rowexport_type_2[ax].remnants_loss_code

      ws.getCell("AA"+count2).value = Number(this.rowexport_type_2[ax].totallenthloss)
      ws.getCell("AA"+count2).numFmt = '0.00';

      ws.getCell("AB"+count2).value = this.rowexport_type_2[ax].totallenthlossper/100
      ws.getCell("AB"+count2).numFmt = '0.00%';
      
      
      this.sum_plug12_t2 = 0
      this.sum_ttl_marker_t2 = Number(this.sum_ttl_marker_t2) + Number(this.rowexport_type_2[ax].ttl_marker)
      this.sum_plug12_t2 = this.rowexport_type_2[ax].plug12 * this.rowexport_type_2[ax].ttl_marker
      this.sum_per_width_loss_t2 = this.rowexport_type_2[ax].widthloss * this.rowexport_type_2[ax].ttl_marker
      this.sum_end1_2_t2 = this.rowexport_type_2[ax].end1_2 * this.rowexport_type_2[ax].ttl_marker
      this.sum_endless_length_yd_t2 = Number(this.sum_endless_length_yd_t2) + Number(this.rowexport_type_2[ax].endless_length_yd)
      this.sum_spliceloss_t2 = Number(this.sum_spliceloss_t2) + Number(this.rowexport_type_2[ax].spliceloss)
      this.sum_totalcutout_t2 = Number(this.sum_totalcutout_t2) + Number(this.rowexport_type_2[ax].totalcutout)
      this.sum_rement_t2 = Number(this.sum_rement_t2) + Number(this.rowexport_type_2[ax].rement)
      this.summary_plug12_t2 = Number(this.summary_plug12_t2) + Number(this.sum_plug12_t2)
      this.summary_widthloss_t2 = Number(this.summary_widthloss_t2) + Number(this.sum_per_width_loss_t2)
      this.summary_end1_2_t2 = Number(this.summary_end1_2_t2) + Number(this.sum_end1_2_t2)
      this.sum_totallenthloss_t2_first = 0
      this.sum_totallenthloss_t2_first = Number(this.rowexport_type_2[ax].endless_length_yd) + Number(this.rowexport_type_2[ax].spliceloss)
      +Number(this.rowexport_type_2[ax].totalcutout) + Number(this.rowexport_type_2[ax].rement)
      this.sum_totallenthloss_t2 = Number(this.sum_totallenthloss_t2) + Number(this.sum_totallenthloss_t2_first)

      this.sum_totallenthloss_per_t2_first = 0
      this.sum_totallenthloss_per_t2_first = Number(this.rowexport_type_2[ax].avg_end) + Number(this.rowexport_type_2[ax].splicelossper)
      +Number(this.rowexport_type_2[ax].totalcutoutper) + Number(this.rowexport_type_2[ax].rement_per)
      this.sum_totallenthloss_per_t2 = Number(this.sum_totallenthloss_t2) + Number(this.sum_totallenthloss_per_t2_first)

      this.sum_number_1_end_t2 = 0 
      this.sum_number_1_end_t2 = this.rowexport_type_2[ax].endless_length_yd * 36
      this.sum_number_2_end_t2 = 0
      this.sum_number_2_end_t2 = this.rowexport_type_2[ax].ttl_marker/this.rowexport_type_2[ax].length_ydb
      this.sum_plug1_2_end_t2 = 0
      this.sum_plug1_2_end_t2 = this.sum_number_1_end_t2/this.sum_number_2_end_t2

      this.sum_plug1_2_end_total_t2 = 0
     
      this.sum_plug1_2_end_total_t2 = this.sum_plug1_2_end_t2 * this.rowexport_type_2[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t2 = Number(this.sum_all_plug1_2_end_total_t2) + Number(this.sum_plug1_2_end_total_t2)

      this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t2
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t2
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t2
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t2
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t2
      })
  
      console.log(this.sum_plug1_2_end_t2)
      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t2,
        value_table:2
      })
      ws.getCell("N"+count2).value = Number(this.sum_plug1_2_end_t2)
      ws.getCell("N"+count2).numFmt = '0.00';
      count2++
         }
         if(this.sum_ttl_marker_t2 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t2
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t2 = this.summary_plug12_t2 / this.sum_ttl_marker_t2
        if(this.sum_plug12_t2 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t2
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t2 =  this.summary_widthloss_t2 / this.sum_ttl_marker_t2 
        if(this.sum_per_width_loss_t2 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t2
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t2 =  this.summary_end1_2_t2/ this.sum_ttl_marker_t2  //M
      if(this.sum_end1_2_t2 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t2
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t2 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t2
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t2 = this.sum_endless_length_yd_t2 / this.sum_ttl_marker_t2 *100 //P
      if(this.sum_avg_end_t2 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t2
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t2 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t2
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t2 = this.sum_spliceloss_t2 / this.sum_ttl_marker_t2 *100
      if(this.sum_splicelossper_t2 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t2
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t2 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t2
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t2 = this.sum_totalcutout_t2 / this.sum_ttl_marker_t2 *100
      if(this.sum_totalcutoutper_t2 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t2
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t2 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t2
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t2 = this.sum_rement_t2 / this.sum_ttl_marker_t2 *100
  if(this.sum_rement_per_t2 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t2
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t2
    if(this.sum_totallenthloss_t2 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t2
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert2 = this.sum_totallenthlosst2 / this.sum_ttl_marker_t2
    if(this.sum_totallenthlosspert2 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert2
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

   ws.getCell("A"+countx2).value = "Total Table A2"  
      ws.mergeCells("A"+countx2+":"+"I"+countx2)
      
      if(this.sum_ttl_marker_t2 > 0 && isNaN(this.sum_ttl_marker_t2)==false && isFinite(this.sum_ttl_marker_t2)==true){
      ws.getCell("J"+countx2).value = Number(this.sum_ttl_marker_t2)
      ws.getCell("J"+countx2).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx2).value = 0
      ws.getCell("J"+countx2).numFmt = '0.00';
      }

      if(this.sum_plug12_t2 > 0 && isNaN(this.sum_plug12_t2)==false && isFinite(this.sum_plug12_t2)==true){
      ws.getCell("K"+countx2).value = Number(this.sum_plug12_t2)
      ws.getCell("K"+countx2).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx2).value = 0
      ws.getCell("K"+countx2).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t2 > 0 && isNaN(this.sum_per_width_loss_t2)==false && isFinite(this.sum_per_width_loss_t2)==true){
      ws.getCell("L"+countx2).value = this.sum_per_width_loss_t2/100
      ws.getCell("L"+countx2).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx2).value = 0.00/100
      ws.getCell("L"+countx2).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t2 = 0
      this.total_result_plug1_2_t2 = this.sum_all_plug1_2_end_total_t2 / this.sum_ttl_marker_t2

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t2
      })
      if(this.total_result_plug1_2_t2 > 0 && isNaN(this.total_result_plug1_2_t2)==false && isFinite(this.total_result_plug1_2_t2)==true){
      ws.getCell("N"+countx2).value = Number(this.total_result_plug1_2_t2)
      ws.getCell("N"+countx2).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx2).value = 0.00
      ws.getCell("N"+countx2).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_t2 > 0 && isNaN(this.sum_endless_length_yd_t2)==false && isFinite(this.sum_endless_length_yd_t2)==true){
      ws.getCell("O"+countx2).value = Number(this.sum_endless_length_yd_t2)
      ws.getCell("O"+countx2).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx2).value = 0.00
      ws.getCell("O"+countx2).numFmt = '0.00';
      }

      if(this.sum_avg_end_t2 > 0 && isNaN(this.sum_avg_end_t2)==false && isFinite(this.sum_avg_end_t2)==true){
      ws.getCell("P"+countx2).value = this.sum_avg_end_t2/100
      ws.getCell("P"+countx2).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx2).value = 0/100
      ws.getCell("P"+countx2).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t2 > 0 && isNaN(this.sum_spliceloss_t2)==false && isFinite(this.sum_spliceloss_t2)==true){
      ws.getCell("R"+countx2).value = Number(this.sum_spliceloss_t2)
      ws.getCell("R"+countx2).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx2).value = 0.00
      ws.getCell("R"+countx2).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t2 > 0 && isNaN(this.sum_splicelossper_t2)==false && isFinite(this.sum_splicelossper_t2)==true){
      ws.getCell("S"+countx2).value = this.sum_splicelossper_t2/100
      ws.getCell("S"+countx2).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx2).value = 0/100
      ws.getCell("S"+countx2).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t2 > 0 && isNaN(this.sum_totalcutout_t2)==false && isFinite(this.sum_totalcutout_t2)==true){
      ws.getCell("U"+countx2).value = Number(this.sum_totalcutout_t2)
      ws.getCell("U"+countx2).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx2).value = 0.00
      ws.getCell("U"+countx2).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t2 > 0 && isNaN(this.sum_totalcutoutper_t2)==false && isFinite(this.sum_totalcutoutper_t2)==true){
      ws.getCell("V"+countx2).value = this.sum_totalcutoutper_t2/100
      ws.getCell("V"+countx2).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx2).value = 0.00/100
      ws.getCell("V"+countx2).numFmt = '0.00%';
      }

       if(this.sum_rement_t2 > 0 && isNaN(this.sum_rement_t2)==false && isFinite(this.sum_rement_t2)==true){
      ws.getCell("X"+countx2).value = Number(this.sum_rement_t2)
      ws.getCell("X"+countx2).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx2).value = 0.00
      ws.getCell("X"+countx2).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t2 > 0 && isNaN(this.sum_rement_per_t2)==false && isFinite(this.sum_rement_per_t2)==true){
      ws.getCell("Y"+countx2).value =  this.sum_rement_per_t2/100
      ws.getCell("Y"+countx2).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx2).value =  0.00/100
      ws.getCell("Y"+countx2).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t2 > 0 && isNaN(this.sum_totallenthloss_t2)==false && isFinite(this.sum_totallenthloss_t2)==true){
      ws.getCell("AA"+countx2).value = Number(this.sum_totallenthloss_t2)
      ws.getCell("AA"+countx2).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx2).value = 0
      ws.getCell("AA"+countx2).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t2 = this.sum_totallenthloss_t2 / this.sum_ttl_marker_t2 *100
      if(this.last_sum_totallenthloss_per_t2 > 0 && isNaN(this.last_sum_totallenthloss_per_t2)==false && isFinite(this.last_sum_totallenthloss_per_t2)==true){
      ws.getCell("AB"+countx2).value = this.last_sum_totallenthloss_per_t2/100
      ws.getCell("AB"+countx2).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx2).value = 0/100
      ws.getCell("AB"+countx2).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx2).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //A3
      var count3 = countx2+1

      if(this.rowexport_type_2.length == 0){
       
         var count3 = count2

       }else{
       var count3 = countx2+1
       }
  
        var countx3 = count3 + 1 + Number(this.rowexport_type_3.length)
      
      this.sum_ttl_marker_t3 = 0
      this.sum_plug12_t3 = 0
      this.sum_per_width_loss_t3 = 0
      this.sum_end1_2_t3 = 0
      this.sum_endless_length_yd_t3= 0
      this.sum_avg_end_t3 = 0
      this.sum_spliceloss_t3 = 0
      this.sum_splicelossper_t3 = 0
      this.sum_totalcutout_t3 = 0
      this.sum_totalcutoutper_t3 = 0
      this.sum_rement_t3 = 0
      this.sum_rement_per_t3 = 0
      this.sum_totallenthloss_t3 =0 
      this.sum_totallenthlossper_t3 = 0
      this.summary_widthloss_t3 = 0
      this.summary_plug12_t3 = 0
      this.summary_end1_2_t3 = 0
      
      
      this.sum_all_plug1_2_end_total_t3 = 0



      if(this.rowexport_type_3.length < 1){
        count_row_grand++
      }else{
          for (var ax = 0; ax<this.rowexport_type_3.length; ax++){ //length +30
      ws.getCell("A"+count3).value = this.rowexport_type_3[ax].mu_date

      ws.getCell("B"+count3).value = this.rowexport_type_3[ax].gm_table

      ws.getCell("C"+count3).value = Number(this.rowexport_type_3[ax].gm_no_short)
      ws.getCell("C"+count3).numFmt = '0.00';

      ws.getCell("D"+count3).value = this.rowexport_type_3[ax].customer

      ws.getCell("E"+count3).value = this.rowexport_type_3[ax].gm_no

      ws.getCell("F"+count3).value = this.rowexport_type_3[ax].fabric_type

      ws.getCell("G"+count3).value = Number(this.rowexport_type_3[ax].obs_width)
      ws.getCell("G"+count3).numFmt = '0.00';

      ws.getCell("H"+count3).value = Number(this.rowexport_type_3[ax].width_inch)
      ws.getCell("H"+count3).numFmt = '0.00';


      ws.getCell("I"+count3).value = Number(this.rowexport_type_3[ax].length_ydb)
      ws.getCell("I"+count3).numFmt = '0.00';

      ws.getCell("J"+count3).value = Number(this.rowexport_type_3[ax].ttl_marker)
      ws.getCell("J"+count3).numFmt = '0.00';

      ws.getCell("K"+count3).value = Number(this.rowexport_type_3[ax].plug12)
      ws.getCell("K"+count3).numFmt = '0.00';

      
      ws.getCell("L"+count3).value = this.rowexport_type_3[ax].widthloss/100
      ws.getCell("L"+count3).numFmt = '0.00%';

      ws.getCell("M"+count3).value = this.rowexport_type_3[ax].widthloss_code
    
      ws.getCell("O"+count3).value = Number(this.rowexport_type_3[ax].endless_length_yd)
      ws.getCell("O"+count3).numFmt = '0.00';

      ws.getCell("P"+count3).value = this.rowexport_type_3[ax].avg_end/100
      ws.getCell("P"+count3).numFmt = '0.00%';

      ws.getCell("Q"+count3).value = this.rowexport_type_3[ax].avg_end_code
      

      ws.getCell("R"+count3).value = Number(this.rowexport_type_3[ax].spliceloss)
      ws.getCell("R"+count3).numFmt = '0.00';

      ws.getCell("S"+count3).value = this.rowexport_type_3[ax].splicelossper/100
      ws.getCell("S"+count3).numFmt = '0.00%';

      ws.getCell("T"+count3).value = this.rowexport_type_3[ax].per_splice_loss_code
   

      ws.getCell("U"+count3).value = Number(this.rowexport_type_3[ax].totalcutout)
      ws.getCell("U"+count3).numFmt = '0.00';

      ws.getCell("V"+count3).value = this.rowexport_type_3[ax].totalcutoutper/100
      ws.getCell("V"+count3).numFmt = '0.00%';

      ws.getCell("W"+count3).value = this.rowexport_type_3[ax].total_cut_out_per_code

      ws.getCell("X"+count3).value = Number(this.rowexport_type_3[ax].rement)
      ws.getCell("X"+count3).numFmt = '0.00';

      ws.getCell("Y"+count3).value = this.rowexport_type_3[ax].rement_per/100
      ws.getCell("Y"+count3).numFmt = '0.00%';

      ws.getCell("Z"+count3).value = this.rowexport_type_3[ax].remnants_loss_code

      ws.getCell("AA"+count3).value = Number(this.rowexport_type_3[ax].totallenthloss)
      ws.getCell("AA"+count3).numFmt = '0.00';

      ws.getCell("AB"+count3).value = this.rowexport_type_3[ax].totallenthlossper/100
      ws.getCell("AB"+count3).numFmt = '0.00%';
      
      this.sum_plug12_t3 = 0
      this.sum_ttl_marker_t3 = Number(this.sum_ttl_marker_t3) + Number(this.rowexport_type_3[ax].ttl_marker)
      this.sum_plug12_t3 = this.rowexport_type_3[ax].plug12 * this.rowexport_type_3[ax].ttl_marker
      this.sum_per_width_loss_t3 = this.rowexport_type_3[ax].widthloss * this.rowexport_type_3[ax].ttl_marker
      this.sum_end1_2_t3 = this.rowexport_type_3[ax].end1_2 * this.rowexport_type_3[ax].ttl_marker
      this.sum_endless_length_yd_t3 = Number(this.sum_endless_length_yd_t3) + Number(this.rowexport_type_3[ax].endless_length_yd)
      this.sum_spliceloss_t3 = Number(this.sum_spliceloss_t3) + Number(this.rowexport_type_3[ax].spliceloss)
      this.sum_totalcutout_t3 = Number(this.sum_totalcutout_t3) + Number(this.rowexport_type_3[ax].totalcutout)
      this.sum_rement_t3 = Number(this.sum_rement_t3) + Number(this.rowexport_type_3[ax].rement)
      this.summary_plug12_t3 = Number(this.summary_plug12_t3) + Number(this.sum_plug12_t3)
      this.summary_widthloss_t3 = Number(this.summary_widthloss_t3) + Number(this.sum_per_width_loss_t3)
      this.summary_end1_2_t3 = Number(this.summary_end1_2_t3) + Number(this.sum_end1_2_t3)
      this.sum_totallenthloss_t3_first = 0
      this.sum_totallenthloss_t3_first = Number(this.rowexport_type_3[ax].endless_length_yd) + Number(this.rowexport_type_3[ax].spliceloss)
      +Number(this.rowexport_type_3[ax].totalcutout) + Number(this.rowexport_type_3[ax].rement)
      this.sum_totallenthloss_t3 = Number(this.sum_totallenthloss_t3) + Number(this.sum_totallenthloss_t3_first)
      
      this.sum_totallenthloss_per_t3_first = 0
      this.sum_totallenthloss_per_t3_first = Number(this.rowexport_type_3[ax].avg_end) + Number(this.rowexport_type_3[ax].splicelossper)
      +Number(this.rowexport_type_3[ax].totalcutoutper) + Number(this.rowexport_type_3[ax].rement_per)
      this.sum_totallenthloss_per_t3 = Number(this.sum_totallenthloss_t3) + Number(this.sum_totallenthloss_per_t3_first)
      this.sum_number_1_end_t3 = 0 
      this.sum_number_1_end_t3 = this.rowexport_type_3[ax].endless_length_yd * 36
      this.sum_number_2_end_t3 = 0
      this.sum_number_2_end_t3 = this.rowexport_type_3[ax].ttl_marker/this.rowexport_type_3[ax].length_ydb
      this.sum_plug1_2_end_t3 = 0
      this.sum_plug1_2_end_t3 = this.sum_number_1_end_t3/this.sum_number_2_end_t3

      this.sum_plug1_2_end_total_t3 = 0
     
      this.sum_plug1_2_end_total_t3 = this.sum_plug1_2_end_t3 * this.rowexport_type_3[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t3 = Number(this.sum_all_plug1_2_end_total_t3) + Number(this.sum_plug1_2_end_total_t3)

      this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t3
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t3
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t3
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t3
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t3
      })

      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t3,
        value_table:3
      })
      ws.getCell("N"+count3).value = Number(this.sum_plug1_2_end_t3)
      ws.getCell("N"+count3).numFmt = '0.00';
      count3++
         }
         if(this.sum_ttl_marker_t3 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t3
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t3 = this.summary_plug12_t3 / this.sum_ttl_marker_t3
        if(this.sum_plug12_t3 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t3
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t3 =  this.summary_widthloss_t3 / this.sum_ttl_marker_t3
        if(this.sum_per_width_loss_t3 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t3
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t3 =  this.summary_end1_2_t3/ this.sum_ttl_marker_t3 //M
      if(this.sum_end1_2_t3 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t3
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t3 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t3
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t3 = this.sum_endless_length_yd_t3 / this.sum_ttl_marker_t3*100 //P
      if(this.sum_avg_end_t3 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t3
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t3 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t3
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t3 = this.sum_spliceloss_t3 / this.sum_ttl_marker_t3*100
      if(this.sum_splicelossper_t3 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t3
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t3 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t3
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t3 = this.sum_totalcutout_t3 / this.sum_ttl_marker_t3*100
      if(this.sum_totalcutoutper_t3 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t3
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t3 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t3
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t3 = this.sum_rement_t3 / this.sum_ttl_marker_t3*100
  if(this.sum_rement_per_t3 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t3
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t3
    if(this.sum_totallenthloss_t3 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t3
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert3 = this.sum_totallenthlosst3 / this.sum_ttl_marker_t3
    if(this.sum_totallenthlosspert3 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert3
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }


     ws.getCell("A"+countx3).value = "Total Table A3"  
      ws.mergeCells("A"+countx3+":"+"I"+countx3)
      
      if(this.sum_ttl_marker_t3 > 0 && isNaN(this.sum_ttl_marker_t3)==false && isFinite(this.sum_ttl_marker_t3)==true){
      ws.getCell("J"+countx3).value = Number(this.sum_ttl_marker_t3)
      ws.getCell("J"+countx3).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx3).value = 0
      ws.getCell("J"+countx3).numFmt = '0.00';
      }

      if(this.sum_plug12_t3 > 0 && isNaN(this.sum_plug12_t3)==false && isFinite(this.sum_plug12_t3)==true){
      ws.getCell("K"+countx3).value = Number(this.sum_plug12_t3)
      ws.getCell("K"+countx3).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx3).value = 0
      ws.getCell("K"+countx3).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t3 > 0 && isNaN(this.sum_per_width_loss_t3)==false && isFinite(this.sum_per_width_loss_t3)==true){
      ws.getCell("L"+countx3).value = this.sum_per_width_loss_t3/100
      ws.getCell("L"+countx3).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx3).value = 0.00/100
      ws.getCell("L"+countx3).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t3 = 0
      this.total_result_plug1_2_t3 = this.sum_all_plug1_2_end_total_t3 / this.sum_ttl_marker_t3

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t3
      })
      if(this.total_result_plug1_2_t3 > 0 && isNaN(this.total_result_plug1_2_t3)==false && isFinite(this.total_result_plug1_2_t3)==true){
      ws.getCell("N"+countx3).value = Number(this.total_result_plug1_2_t3)
      ws.getCell("N"+countx3).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx3).value = 0.00
      ws.getCell("N"+countx3).numFmt = '0.00';
      }
      

      if(this.sum_endless_length_yd_t3 > 0 && isNaN(this.sum_endless_length_yd_t3)==false && isFinite(this.sum_endless_length_yd_t3)==true){
      ws.getCell("O"+countx3).value = Number(this.sum_endless_length_yd_t3)
      ws.getCell("O"+countx3).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx3).value = 0.00
      ws.getCell("O"+countx3).numFmt = '0.00';
      }

      if(this.sum_avg_end_t3 > 0 && isNaN(this.sum_avg_end_t3)==false && isFinite(this.sum_avg_end_t3)==true){
      ws.getCell("P"+countx3).value = this.sum_avg_end_t3/100
      ws.getCell("P"+countx3).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx3).value = 0/100
      ws.getCell("P"+countx3).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t3 > 0 && isNaN(this.sum_spliceloss_t3)==false && isFinite(this.sum_spliceloss_t3)==true){
      ws.getCell("R"+countx3).value = Number(this.sum_spliceloss_t3)
      ws.getCell("R"+countx3).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx3).value = 0.00
      ws.getCell("R"+countx3).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t3 > 0 && isNaN(this.sum_splicelossper_t3)==false && isFinite(this.sum_splicelossper_t3)==true){
      ws.getCell("S"+countx3).value = this.sum_splicelossper_t3/100
      ws.getCell("S"+countx3).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx3).value = 0/100
      ws.getCell("S"+countx3).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t3 > 0 && isNaN(this.sum_totalcutout_t3)==false && isFinite(this.sum_totalcutout_t3)==true){
      ws.getCell("U"+countx3).value = Number(this.sum_totalcutout_t3)
      ws.getCell("U"+countx3).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx3).value = 0.00
      ws.getCell("U"+countx3).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t3 > 0 && isNaN(this.sum_totalcutoutper_t3)==false && isFinite(this.sum_totalcutoutper_t3)==true){
      ws.getCell("V"+countx3).value = this.sum_totalcutoutper_t3/100
      ws.getCell("V"+countx3).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx3).value = 0.00/100
      ws.getCell("V"+countx3).numFmt = '0.00%';
      }

       if(this.sum_rement_t3 > 0 && isNaN(this.sum_rement_t3)==false && isFinite(this.sum_rement_t3)==true){
      ws.getCell("X"+countx3).value = Number(this.sum_rement_t3)
      ws.getCell("X"+countx3).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx3).value = 0.00
      ws.getCell("X"+countx3).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t3 > 0 && isNaN(this.sum_rement_per_t3)==false && isFinite(this.sum_rement_per_t3)==true){
      ws.getCell("Y"+countx3).value =  this.sum_rement_per_t3/100
      ws.getCell("Y"+countx3).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx3).value =  0.00/100
      ws.getCell("Y"+countx3).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t3 > 0 && isNaN(this.sum_totallenthloss_t3)==false && isFinite(this.sum_totallenthloss_t3)==true){
      ws.getCell("AA"+countx3).value = Number(this.sum_totallenthloss_t3)
      ws.getCell("AA"+countx3).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx3).value = 0
      ws.getCell("AA"+countx3).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t3 = this.sum_totallenthloss_t3 / this.sum_ttl_marker_t3 *100
      if(this.last_sum_totallenthloss_per_t3 > 0 && isNaN(this.last_sum_totallenthloss_per_t3)==false && isFinite(this.last_sum_totallenthloss_per_t3)==true){
      ws.getCell("AB"+countx3).value = this.last_sum_totallenthloss_per_t3/100
      ws.getCell("AB"+countx3).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx3).value = 0/100
      ws.getCell("AB"+countx3).numFmt = '0.00%'; 
      }


      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx3).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //A4
       
   

   if(this.rowexport_type_3.length == 0){
       
         var count4 = count3

 
       }else{
      var count4 = countx3+1
       }
        var countx4 = count4 + 1 + Number(this.rowexport_type_4.length)
      
    
      this.sum_ttl_marker_t4 = 0
      this.sum_plug12_t4 = 0
      this.sum_per_width_loss_t4 = 0
      this.sum_end1_2_t4 = 0
      this.sum_endless_length_yd_t4= 0
      this.sum_avg_end_t4 = 0
      this.sum_spliceloss_t4 = 0
      this.sum_splicelossper_t4 = 0
      this.sum_totalcutout_t4 = 0
      this.sum_totalcutoutper_t4 = 0
      this.sum_rement_t4 = 0
      this.sum_rement_per_t4 = 0
      this.sum_totallenthloss_t4 =0 
      this.sum_totallenthlossper_t4 = 0
      this.summary_widthloss_t4 = 0
      this.summary_plug12_t4 = 0
      this.summary_end1_2_t4 = 0

    
      
      this.sum_all_plug1_2_end_total_t4 = 0


      if(this.rowexport_type_4.length < 1){
        count_row_grand++
      }else{
          for (var ax = 0; ax<this.rowexport_type_4.length; ax++){ //length +40
      ws.getCell("A"+count4).value = this.rowexport_type_4[ax].mu_date

      ws.getCell("B"+count4).value = this.rowexport_type_4[ax].gm_table

      ws.getCell("C"+count4).value = Number(this.rowexport_type_4[ax].gm_no_short)
      ws.getCell("C"+count4).numFmt = '0.00';

      ws.getCell("D"+count4).value = this.rowexport_type_4[ax].customer

      ws.getCell("E"+count4).value = this.rowexport_type_4[ax].gm_no

      ws.getCell("F"+count4).value = this.rowexport_type_4[ax].fabric_type

      ws.getCell("G"+count4).value = Number(this.rowexport_type_4[ax].obs_width)
      ws.getCell("G"+count4).numFmt = '0.00';

      ws.getCell("H"+count4).value = Number(this.rowexport_type_4[ax].width_inch)
      ws.getCell("H"+count4).numFmt = '0.00';


      ws.getCell("I"+count4).value = Number(this.rowexport_type_4[ax].length_ydb)
      ws.getCell("I"+count4).numFmt = '0.00';

      ws.getCell("J"+count4).value = Number(this.rowexport_type_4[ax].ttl_marker)
      ws.getCell("J"+count4).numFmt = '0.00';

      ws.getCell("K"+count4).value = Number(this.rowexport_type_4[ax].plug12)
      ws.getCell("K"+count4).numFmt = '0.00';

      ws.getCell("L"+count4).value = this.rowexport_type_4[ax].widthloss/100
      ws.getCell("L"+count4).numFmt = '0.00%';

      ws.getCell("M"+count4).value = this.rowexport_type_4[ax].widthloss_code
   
      ws.getCell("O"+count4).value = Number(this.rowexport_type_4[ax].endless_length_yd)
      ws.getCell("O"+count4).numFmt = '0.00';

      ws.getCell("P"+count4).value = this.rowexport_type_4[ax].avg_end/100
      ws.getCell("P"+count4).numFmt = '0.00%';

      ws.getCell("Q"+count4).value = this.rowexport_type_4[ax].avg_end_code
      

      ws.getCell("R"+count4).value = Number(this.rowexport_type_4[ax].spliceloss)
      ws.getCell("R"+count4).numFmt = '0.00';

      ws.getCell("S"+count4).value = this.rowexport_type_4[ax].splicelossper/100
      ws.getCell("S"+count4).numFmt = '0.00%';

      ws.getCell("T"+count4).value = this.rowexport_type_4[ax].per_splice_loss_code
   

      ws.getCell("U"+count4).value = Number(this.rowexport_type_4[ax].totalcutout)
      ws.getCell("U"+count4).numFmt = '0.00';

      ws.getCell("V"+count4).value = this.rowexport_type_4[ax].totalcutoutper/100
      ws.getCell("V"+count4).numFmt = '0.00%';

      ws.getCell("W"+count4).value = this.rowexport_type_4[ax].total_cut_out_per_code

      ws.getCell("X"+count4).value = Number(this.rowexport_type_4[ax].rement)
      ws.getCell("X"+count4).numFmt = '0.00';

      ws.getCell("Y"+count4).value = this.rowexport_type_4[ax].rement_per/100
      ws.getCell("Y"+count4).numFmt = '0.00%';

      ws.getCell("Z"+count4).value = this.rowexport_type_4[ax].remnants_loss_code

      ws.getCell("AA"+count4).value = Number(this.rowexport_type_4[ax].totallenthloss)
      ws.getCell("AA"+count4).numFmt = '0.00';

      ws.getCell("AB"+count4).value = this.rowexport_type_4[ax].totallenthlossper/100
      ws.getCell("AB"+count4).numFmt = '0.00%';
      
      this.sum_plug12_t4 = 0
      this.sum_ttl_marker_t4 = Number(this.sum_ttl_marker_t4) + Number(this.rowexport_type_4[ax].ttl_marker)
      this.sum_plug12_t4 = this.rowexport_type_4[ax].plug12 * this.rowexport_type_4[ax].ttl_marker
      this.sum_per_width_loss_t4 = this.rowexport_type_4[ax].widthloss * this.rowexport_type_4[ax].ttl_marker
      this.sum_end1_2_t4 = this.rowexport_type_4[ax].end1_2 * this.rowexport_type_4[ax].ttl_marker
      this.sum_endless_length_yd_t4 = Number(this.sum_endless_length_yd_t4) + Number(this.rowexport_type_4[ax].endless_length_yd)
      this.sum_spliceloss_t4 = Number(this.sum_spliceloss_t4) + Number(this.rowexport_type_4[ax].spliceloss)
      this.sum_totalcutout_t4 = Number(this.sum_totalcutout_t4) + Number(this.rowexport_type_4[ax].totalcutout)
      this.sum_rement_t4 = Number(this.sum_rement_t4) + Number(this.rowexport_type_4[ax].rement)
        this.summary_plug12_t4 = Number(this.summary_plug12_t4) + Number(this.sum_plug12_t4)
      this.summary_widthloss_t4 = Number(this.summary_widthloss_t4) + Number(this.sum_per_width_loss_t4)
      this.summary_end1_2_t4 = Number(this.summary_end1_2_t4) + Number(this.sum_end1_2_t4)
     
      this.sum_totallenthloss_t4_first = 0
      this.sum_totallenthloss_t4_first = Number(this.rowexport_type_4[ax].endless_length_yd) + Number(this.rowexport_type_4[ax].spliceloss)
      +Number(this.rowexport_type_4[ax].totalcutout) + Number(this.rowexport_type_4[ax].rement)
      this.sum_totallenthloss_t4 = Number(this.sum_totallenthloss_t4) + Number(this.sum_totallenthloss_t4_first)

      this.sum_totallenthloss_per_t4_first = 0
      this.sum_totallenthloss_per_t4_first = Number(this.rowexport_type_4[ax].avg_end) + Number(this.rowexport_type_4[ax].splicelossper)
      +Number(this.rowexport_type_4[ax].totalcutoutper) + Number(this.rowexport_type_4[ax].rement_per)
      this.sum_totallenthloss_per_t4 = Number(this.sum_totallenthloss_t4) + Number(this.sum_totallenthloss_per_t4_first)
      this.sum_number_1_end_t4 = 0 
      this.sum_number_1_end_t4 = this.rowexport_type_4[ax].endless_length_yd * 36
      this.sum_number_2_end_t4 = 0
      this.sum_number_2_end_t4 = this.rowexport_type_4[ax].ttl_marker/this.rowexport_type_4[ax].length_ydb
      this.sum_plug1_2_end_t4 = 0
      this.sum_plug1_2_end_t4 = this.sum_number_1_end_t4/this.sum_number_2_end_t4

      this.sum_plug1_2_end_total_t4 = 0
     
      this.sum_plug1_2_end_total_t4 = this.sum_plug1_2_end_t4 * this.rowexport_type_4[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t4 = Number(this.sum_all_plug1_2_end_total_t4) + Number(this.sum_plug1_2_end_total_t4)

      this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t4
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t4
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t4
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t4
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t4
      })

      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t4,
        value_table:4
      })

      ws.getCell("N"+count4).value = Number(this.sum_plug1_2_end_t4)
      ws.getCell("N"+count4).numFmt = '0.00';

      count4++
         }
         if(this.sum_ttl_marker_t4 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t4
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t4 = this.summary_plug12_t4 / this.sum_ttl_marker_t4
        if(this.sum_plug12_t4 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t4
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t4 =  this.summary_widthloss_t4 / this.sum_ttl_marker_t4
        if(this.sum_per_width_loss_t4 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t4
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t4 =  this.summary_end1_2_t4/ this.sum_ttl_marker_t4 //M
      if(this.sum_end1_2_t4 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t4
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t4 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t4
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t4 = this.sum_endless_length_yd_t4 / this.sum_ttl_marker_t4*100 //P
      if(this.sum_avg_end_t4 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t4
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t4 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t4
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t4 = this.sum_spliceloss_t4 / this.sum_ttl_marker_t4*100
      if(this.sum_splicelossper_t4 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t4
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t4 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t4
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t4 = this.sum_totalcutout_t4 / this.sum_ttl_marker_t4*100
      if(this.sum_totalcutoutper_t4 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t4
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t4 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t4
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t4 = this.sum_rement_t4 / this.sum_ttl_marker_t4*100
  if(this.sum_rement_per_t4 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t4
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t4
    if(this.sum_totallenthloss_t4 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t4
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert4 = this.sum_totallenthlosst4 / this.sum_ttl_marker_t4
    if(this.sum_totallenthlosspert4 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert4
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

ws.getCell("A"+countx4).value = "Total Table A4"  
      ws.mergeCells("A"+countx4+":"+"I"+countx4)
      
      if(this.sum_ttl_marker_t4 > 0 && isNaN(this.sum_ttl_marker_t4)==false && isFinite(this.sum_ttl_marker_t4)==true){
      ws.getCell("J"+countx4).value = Number(this.sum_ttl_marker_t4)
      ws.getCell("J"+countx4).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx4).value = 0
      ws.getCell("J"+countx4).numFmt = '0.00';
      }

      if(this.sum_plug12_t4 > 0 && isNaN(this.sum_plug12_t4)==false && isFinite(this.sum_plug12_t4)==true){
      ws.getCell("K"+countx4).value = Number(this.sum_plug12_t4)
      ws.getCell("K"+countx4).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx4).value = 0
      ws.getCell("K"+countx4).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t4 > 0 && isNaN(this.sum_per_width_loss_t4)==false && isFinite(this.sum_per_width_loss_t4)==true){
      ws.getCell("L"+countx4).value = this.sum_per_width_loss_t4/100
      ws.getCell("L"+countx4).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx4).value = 0.00/100
      ws.getCell("L"+countx4).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t4 = 0
      this.total_result_plug1_2_t4 = this.sum_all_plug1_2_end_total_t4 / this.sum_ttl_marker_t4

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t4
      })
      if(this.total_result_plug1_2_t4 > 0 && isNaN(this.total_result_plug1_2_t4)==false && isFinite(this.total_result_plug1_2_t4)==true){
      ws.getCell("N"+countx4).value = Number(this.total_result_plug1_2_t4)
      ws.getCell("N"+countx4).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx4).value = 0.00
      ws.getCell("N"+countx4).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_t4 > 0 && isNaN(this.sum_endless_length_yd_t4)==false && isFinite(this.sum_endless_length_yd_t4)==true){
      ws.getCell("O"+countx4).value = Number(this.sum_endless_length_yd_t4)
      ws.getCell("O"+countx4).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx4).value = 0.00
      ws.getCell("O"+countx4).numFmt = '0.00';
      }

      if(this.sum_avg_end_t4 > 0 && isNaN(this.sum_avg_end_t4)==false && isFinite(this.sum_avg_end_t4)==true){
      ws.getCell("P"+countx4).value = this.sum_avg_end_t4/100
      ws.getCell("P"+countx4).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx4).value = 0/100
      ws.getCell("P"+countx4).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t4 > 0 && isNaN(this.sum_spliceloss_t4)==false && isFinite(this.sum_spliceloss_t4)==true){
      ws.getCell("R"+countx4).value = Number(this.sum_spliceloss_t4)
      ws.getCell("R"+countx4).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx4).value = 0.00
      ws.getCell("R"+countx4).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t4 > 0 && isNaN(this.sum_splicelossper_t4)==false && isFinite(this.sum_splicelossper_t4)==true){
      ws.getCell("S"+countx4).value = this.sum_splicelossper_t4/100
      ws.getCell("S"+countx4).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx4).value = 0/100
      ws.getCell("S"+countx4).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t4 > 0 && isNaN(this.sum_totalcutout_t4)==false && isFinite(this.sum_totalcutout_t4)==true){
      ws.getCell("U"+countx4).value = Number(this.sum_totalcutout_t4)
      ws.getCell("U"+countx4).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx4).value = 0.00
      ws.getCell("U"+countx4).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t4 > 0 && isNaN(this.sum_totalcutoutper_t4)==false && isFinite(this.sum_totalcutoutper_t4)==true){
      ws.getCell("V"+countx4).value = this.sum_totalcutoutper_t4/100
      ws.getCell("V"+countx4).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx4).value = 0.00/100
      ws.getCell("V"+countx4).numFmt = '0.00%';
      }

       if(this.sum_rement_t4 > 0 && isNaN(this.sum_rement_t4)==false && isFinite(this.sum_rement_t4)==true){
      ws.getCell("X"+countx4).value = Number(this.sum_rement_t4)
      ws.getCell("X"+countx4).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx4).value = 0.00
      ws.getCell("X"+countx4).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t4 > 0 && isNaN(this.sum_rement_per_t4)==false && isFinite(this.sum_rement_per_t4)==true){
      ws.getCell("Y"+countx4).value =  this.sum_rement_per_t4/100
      ws.getCell("Y"+countx4).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx4).value =  0.00/100
      ws.getCell("Y"+countx4).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t4 > 0 && isNaN(this.sum_totallenthloss_t4)==false && isFinite(this.sum_totallenthloss_t4)==true){
      ws.getCell("AA"+countx4).value = Number(this.sum_totallenthloss_t4)
      ws.getCell("AA"+countx4).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx4).value = 0
      ws.getCell("AA"+countx4).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t4 = this.sum_totallenthloss_t4 / this.sum_ttl_marker_t4 *100
      if(this.last_sum_totallenthloss_per_t4 > 0 && isNaN(this.last_sum_totallenthloss_per_t4)==false && isFinite(this.last_sum_totallenthloss_per_t4)==true){
      ws.getCell("AB"+countx4).value = this.last_sum_totallenthloss_per_t4/100
      ws.getCell("AB"+countx4).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx4).value = 0/100
      ws.getCell("AB"+countx4).numFmt = '0.00%'; 
      }

      
      
      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx4).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //A5
       
   if(this.rowexport_type_4.length == 0){
       
         var count5 = count4

 
       }else{
      var count5 = countx4+1
       }
        var countx5 = count5 + 1 +Number(this.rowexport_type_5.length)
      
      this.sum_ttl_marker_t5 = 0
      this.sum_plug12_t5 = 0
      this.sum_per_width_loss_t5 = 0
      this.sum_end1_2_t5 = 0
      this.sum_endless_length_yd_t5= 0
      this.sum_avg_end_t5 = 0
      this.sum_spliceloss_t5 = 0
      this.sum_splicelossper_t5 = 0
      this.sum_totalcutout_t5 = 0
      this.sum_totalcutoutper_t5 = 0
      this.sum_rement_t5 = 0
      this.sum_rement_per_t5 = 0
      this.sum_totallenthloss_t5 =0 
      this.sum_totallenthlossper_t5 = 0
      this.summary_widthloss_t5 = 0
      this.summary_plug12_t5 = 0
      this.summary_end1_2_t5 = 0

      
     this.sum_all_plug1_2_end_total_t5 = 0

     if(this.rowexport_type_5.length < 1){
        count_row_grand++
      }else{
      
          for (var ax = 0; ax<this.rowexport_type_5.length; ax++){ //length +50
      ws.getCell("A"+count5).value = this.rowexport_type_5[ax].mu_date

      ws.getCell("B"+count5).value = this.rowexport_type_5[ax].gm_table

      ws.getCell("C"+count5).value = Number(this.rowexport_type_5[ax].gm_no_short)
      ws.getCell("C"+count5).numFmt = '0.00';

      ws.getCell("D"+count5).value = this.rowexport_type_5[ax].customer

      ws.getCell("E"+count5).value = this.rowexport_type_5[ax].gm_no

      ws.getCell("F"+count5).value = this.rowexport_type_5[ax].fabric_type

      ws.getCell("G"+count5).value = Number(this.rowexport_type_5[ax].obs_width)
      ws.getCell("G"+count5).numFmt = '0.00';

      ws.getCell("H"+count5).value = Number(this.rowexport_type_5[ax].width_inch)
      ws.getCell("H"+count5).numFmt = '0.00';


      ws.getCell("I"+count5).value = Number(this.rowexport_type_5[ax].length_ydb)
      ws.getCell("I"+count5).numFmt = '0.00';

      ws.getCell("J"+count5).value = Number(this.rowexport_type_5[ax].ttl_marker)
      ws.getCell("J"+count5).numFmt = '0.00';

      ws.getCell("K"+count5).value = Number(this.rowexport_type_5[ax].plug12)
      ws.getCell("K"+count5).numFmt = '0.00';

      ws.getCell("L"+count5).value = this.rowexport_type_5[ax].widthloss/100
      ws.getCell("L"+count5).numFmt = '0.00%';

      ws.getCell("M"+count5).value = this.rowexport_type_5[ax].widthloss_code
    
      ws.getCell("O"+count5).value = Number(this.rowexport_type_5[ax].endless_length_yd)
      ws.getCell("O"+count5).numFmt = '0.00';

      ws.getCell("P"+count5).value = this.rowexport_type_5[ax].avg_end/100
      ws.getCell("P"+count5).numFmt = '0.00%';

      ws.getCell("Q"+count5).value = this.rowexport_type_5[ax].avg_end_code
      

      ws.getCell("R"+count5).value = Number(this.rowexport_type_5[ax].spliceloss)
      ws.getCell("R"+count5).numFmt = '0.00';

      ws.getCell("S"+count5).value = this.rowexport_type_5[ax].splicelossper/100
      ws.getCell("S"+count5).numFmt = '0.00%';

      ws.getCell("T"+count5).value = this.rowexport_type_5[ax].per_splice_loss_code
   

      ws.getCell("U"+count5).value = Number(this.rowexport_type_5[ax].totalcutout)
      ws.getCell("U"+count5).numFmt = '0.00';

      ws.getCell("V"+count5).value = this.rowexport_type_5[ax].totalcutoutper/100
      ws.getCell("V"+count5).numFmt = '0.00%';

      ws.getCell("W"+count5).value = this.rowexport_type_5[ax].total_cut_out_per_code

      ws.getCell("X"+count5).value = Number(this.rowexport_type_5[ax].rement)
      ws.getCell("X"+count5).numFmt = '0.00';

      ws.getCell("Y"+count5).value = this.rowexport_type_5[ax].rement_per/100
      ws.getCell("Y"+count5).numFmt = '0.00%';

      ws.getCell("Z"+count5).value = this.rowexport_type_5[ax].remnants_loss_code

      ws.getCell("AA"+count5).value = Number(this.rowexport_type_5[ax].totallenthloss)
      ws.getCell("AA"+count5).numFmt = '0.00';

      ws.getCell("AB"+count5).value = this.rowexport_type_5[ax].totallenthlossper/100
      ws.getCell("AB"+count5).numFmt = '0.00%';
      
      this.sum_plug12_t5 = 0
      this.sum_ttl_marker_t5 = Number(this.sum_ttl_marker_t5) + Number(this.rowexport_type_5[ax].ttl_marker)
      this.sum_plug12_t5 = this.rowexport_type_5[ax].plug12 * this.rowexport_type_5[ax].ttl_marker
      this.sum_per_width_loss_t5 = this.rowexport_type_5[ax].widthloss * this.rowexport_type_5[ax].ttl_marker
      this.sum_end1_2_t5 = this.rowexport_type_5[ax].end1_2 * this.rowexport_type_5[ax].ttl_marker
      this.sum_endless_length_yd_t5 = Number(this.sum_endless_length_yd_t5) + Number(this.rowexport_type_5[ax].endless_length_yd)
      this.sum_spliceloss_t5 = Number(this.sum_spliceloss_t5) + Number(this.rowexport_type_5[ax].spliceloss)
      this.sum_totalcutout_t5 = Number(this.sum_totalcutout_t5) + Number(this.rowexport_type_5[ax].totalcutout)
      this.sum_rement_t5 = Number(this.sum_rement_t5) + Number(this.rowexport_type_5[ax].rement)
      this.summary_plug12_t5 = Number(this.summary_plug12_t5) + Number(this.sum_plug12_t5)
      this.summary_widthloss_t5 = Number(this.summary_widthloss_t5) + Number(this.sum_per_width_loss_t5)
      this.summary_end1_2_t5 = Number(this.summary_end1_2_t5) + Number(this.sum_end1_2_t5)
      this.sum_totallenthloss_t5_first = 0
      this.sum_totallenthloss_t5_first = Number(this.rowexport_type_5[ax].endless_length_yd) + Number(this.rowexport_type_5[ax].spliceloss)
      +Number(this.rowexport_type_5[ax].totalcutout) + Number(this.rowexport_type_5[ax].rement)
      this.sum_totallenthloss_t5 = Number(this.sum_totallenthloss_t5) + Number(this.sum_totallenthloss_t5_first)

      this.sum_totallenthloss_per_t5_first = 0
      this.sum_totallenthloss_per_t5_first = Number(this.rowexport_type_5[ax].avg_end) + Number(this.rowexport_type_5[ax].splicelossper)
      +Number(this.rowexport_type_5[ax].totalcutoutper) + Number(this.rowexport_type_5[ax].rement_per)
      this.sum_number_1_end_t5 = 0 
      this.sum_number_1_end_t5 = this.rowexport_type_5[ax].endless_length_yd * 36
      this.sum_number_2_end_t5 = 0
      this.sum_number_2_end_t5 = this.rowexport_type_5[ax].ttl_marker/this.rowexport_type_5[ax].length_ydb
      this.sum_plug1_2_end_t5 = 0
      this.sum_plug1_2_end_t5 = this.sum_number_1_end_t5/this.sum_number_2_end_t5

      this.sum_plug1_2_end_total_t5 = 0
     
      this.sum_plug1_2_end_total_t5 = this.sum_plug1_2_end_t5 * this.rowexport_type_5[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t5 = Number(this.sum_all_plug1_2_end_total_t5) + Number(this.sum_plug1_2_end_total_t5)

      this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t5
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t5
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t5
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t5
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t5
      })

      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t5,
        value_table:5
      })
       ws.getCell("N"+count5).value = Number(this.sum_plug1_2_end_t5)
      ws.getCell("N"+count5).numFmt = '0.00';
      count5++
         }
         if(this.sum_ttl_marker_t5 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t5
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t5 = this.summary_plug12_t5  / this.sum_ttl_marker_t5
        if(this.sum_plug12_t5 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t5
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t5 =  this.summary_widthloss_t5 / this.sum_ttl_marker_t5
        if(this.sum_per_width_loss_t5 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t5
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t5 =  this.summary_end1_2_t5/ this.sum_ttl_marker_t5 //M
      if(this.sum_end1_2_t5 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t5
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t5 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t5
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t5 = this.sum_endless_length_yd_t5 / this.sum_ttl_marker_t5*100 //P
      if(this.sum_avg_end_t5 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t5
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t5 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t5
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t5 = this.sum_spliceloss_t5 / this.sum_ttl_marker_t5*100
      if(this.sum_splicelossper_t5 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t5
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t5 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t5
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t5 = this.sum_totalcutout_t5 / this.sum_ttl_marker_t5*100
      if(this.sum_totalcutoutper_t5 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t5
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t5 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t5
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t5 = this.sum_rement_t5 / this.sum_ttl_marker_t5*100
  if(this.sum_rement_per_t5 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t5
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t5
    if(this.sum_totallenthloss_t5 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t5
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert5 = this.sum_totallenthlosst5 / this.sum_ttl_marker_t5
    if(this.sum_totallenthlosspert5 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert5
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }


      ws.getCell("A"+countx5).value = "Total Table A5"  
      ws.mergeCells("A"+countx5+":"+"I"+countx5)
      
      if(this.sum_ttl_marker_t5 > 0 && isNaN(this.sum_ttl_marker_t5)==false && isFinite(this.sum_ttl_marker_t5)==true){
      ws.getCell("J"+countx5).value = Number(this.sum_ttl_marker_t5)
      ws.getCell("J"+countx5).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx5).value = 0
      ws.getCell("J"+countx5).numFmt = '0.00';
      }

      if(this.sum_plug12_t5 > 0 && isNaN(this.sum_plug12_t5)==false && isFinite(this.sum_plug12_t5)==true){
      ws.getCell("K"+countx5).value = Number(this.sum_plug12_t5)
      ws.getCell("K"+countx5).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx5).value = 0
      ws.getCell("K"+countx5).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t5 > 0 && isNaN(this.sum_per_width_loss_t5)==false && isFinite(this.sum_per_width_loss_t5)==true){
      ws.getCell("L"+countx5).value = this.sum_per_width_loss_t5/100
      ws.getCell("L"+countx5).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx5).value = 0.00/100
      ws.getCell("L"+countx5).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t5 = 0
      this.total_result_plug1_2_t5 = this.sum_all_plug1_2_end_total_t5 / this.sum_ttl_marker_t5

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t5
      })
      if(this.total_result_plug1_2_t5 > 0 && isNaN(this.total_result_plug1_2_t5)==false && isFinite(this.total_result_plug1_2_t5)==true){
      ws.getCell("N"+countx5).value = Number(this.total_result_plug1_2_t5)
      ws.getCell("N"+countx5).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx5).value = 0.00
      ws.getCell("N"+countx5).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_t5 > 0 && isNaN(this.sum_endless_length_yd_t5)==false && isFinite(this.sum_endless_length_yd_t5)==true){
      ws.getCell("O"+countx5).value = Number(this.sum_endless_length_yd_t5)
      ws.getCell("O"+countx5).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx5).value = 0.00
      ws.getCell("O"+countx5).numFmt = '0.00';
      }

      if(this.sum_avg_end_t5 > 0 && isNaN(this.sum_avg_end_t5)==false && isFinite(this.sum_avg_end_t5)==true){
      ws.getCell("P"+countx5).value = this.sum_avg_end_t5/100
      ws.getCell("P"+countx5).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx5).value = 0/100
      ws.getCell("P"+countx5).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t5 > 0 && isNaN(this.sum_spliceloss_t5)==false && isFinite(this.sum_spliceloss_t5)==true){
      ws.getCell("R"+countx5).value = Number(this.sum_spliceloss_t5)
      ws.getCell("R"+countx5).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx5).value = 0.00
      ws.getCell("R"+countx5).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t5 > 0 && isNaN(this.sum_splicelossper_t5)==false && isFinite(this.sum_splicelossper_t5)==true){
      ws.getCell("S"+countx5).value = this.sum_splicelossper_t5/100
      ws.getCell("S"+countx5).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx5).value = 0/100
      ws.getCell("S"+countx5).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t5 > 0 && isNaN(this.sum_totalcutout_t5)==false && isFinite(this.sum_totalcutout_t5)==true){
      ws.getCell("U"+countx5).value = Number(this.sum_totalcutout_t5)
      ws.getCell("U"+countx5).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx5).value = 0.00
      ws.getCell("U"+countx5).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t5 > 0 && isNaN(this.sum_totalcutoutper_t5)==false && isFinite(this.sum_totalcutoutper_t5)==true){
      ws.getCell("V"+countx5).value = this.sum_totalcutoutper_t5/100
      ws.getCell("V"+countx5).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx5).value = 0.00/100
      ws.getCell("V"+countx5).numFmt = '0.00%';
      }

       if(this.sum_rement_t5 > 0 && isNaN(this.sum_rement_t5)==false && isFinite(this.sum_rement_t5)==true){
      ws.getCell("X"+countx5).value = Number(this.sum_rement_t5)
      ws.getCell("X"+countx5).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx5).value = 0.00
      ws.getCell("X"+countx5).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t5 > 0 && isNaN(this.sum_rement_per_t5)==false && isFinite(this.sum_rement_per_t5)==true){
      ws.getCell("Y"+countx5).value =  this.sum_rement_per_t5/100
      ws.getCell("Y"+countx5).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx5).value =  0.00/100
      ws.getCell("Y"+countx5).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t5 > 0 && isNaN(this.sum_totallenthloss_t5)==false && isFinite(this.sum_totallenthloss_t5)==true){
      ws.getCell("AA"+countx5).value = Number(this.sum_totallenthloss_t5)
      ws.getCell("AA"+countx5).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx5).value = 0
      ws.getCell("AA"+countx5).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t5 = this.sum_totallenthloss_t5 / this.sum_ttl_marker_t5 *100
      if(this.last_sum_totallenthloss_per_t5 > 0 && isNaN(this.last_sum_totallenthloss_per_t5)==false && isFinite(this.last_sum_totallenthloss_per_t5)==true){
      ws.getCell("AB"+countx5).value = this.last_sum_totallenthloss_per_t5/100
      ws.getCell("AB"+countx5).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx5).value = 0/100
      ws.getCell("AB"+countx5).numFmt = '0.00%'; 
      }
      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx5).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //A6
       if(this.rowexport_type_5.length == 0){
         var count6 = count5
       }else{
       var count6 = countx5+1
       }
 
        var countx6 = count6 + 1 +Number(this.rowexport_type_6.length)
      
      this.sum_ttl_marker_t6 = 0
      this.sum_plug12_t6 = 0
      this.sum_per_width_loss_t6 = 0
      this.sum_end1_2_t6 = 0
      this.sum_endless_length_yd_t6= 0
      this.sum_avg_end_t6 = 0
      this.sum_spliceloss_t6 = 0
      this.sum_splicelossper_t6 = 0
      this.sum_totalcutout_t6 = 0
      this.sum_totalcutoutper_t6 = 0
      this.sum_rement_t6 = 0
      this.sum_rement_per_t6 = 0
      this.sum_totallenthloss_t6 =0 
      this.sum_totallenthlossper_t6 = 0
      this.summary_widthloss_t6 = 0
      this.summary_plug12_t6 = 0
      this.summary_end1_2_t6 = 0

      
     if(this.rowexport_type_6.length < 1){
        count_row_grand++
      }else{
      this.sum_all_plug1_2_end_total_t6 = 0
          for (var ax = 0; ax<this.rowexport_type_6.length; ax++){ //length +60
      ws.getCell("A"+count6).value = this.rowexport_type_6[ax].mu_date

      ws.getCell("B"+count6).value = this.rowexport_type_6[ax].gm_table

      ws.getCell("C"+count6).value = Number(this.rowexport_type_6[ax].gm_no_short)
      ws.getCell("C"+count6).numFmt = '0.00';

      ws.getCell("D"+count6).value = this.rowexport_type_6[ax].customer

      ws.getCell("E"+count6).value = this.rowexport_type_6[ax].gm_no

      ws.getCell("F"+count6).value = this.rowexport_type_6[ax].fabric_type

      ws.getCell("G"+count6).value = Number(this.rowexport_type_6[ax].obs_width)
      ws.getCell("G"+count6).numFmt = '0.00';

      ws.getCell("H"+count6).value = Number(this.rowexport_type_6[ax].width_inch)
      ws.getCell("H"+count6).numFmt = '0.00';


      ws.getCell("I"+count6).value = Number(this.rowexport_type_6[ax].length_ydb)
      ws.getCell("I"+count6).numFmt = '0.00';

      ws.getCell("J"+count6).value = Number(this.rowexport_type_6[ax].ttl_marker)
      ws.getCell("J"+count6).numFmt = '0.00';

      ws.getCell("K"+count6).value = Number(this.rowexport_type_6[ax].plug12)
      ws.getCell("K"+count6).numFmt = '0.00';

      ws.getCell("L"+count6).value = this.rowexport_type_6[ax].widthloss/100
      ws.getCell("L"+count6).numFmt = '0.00%';

      ws.getCell("M"+count6).value = this.rowexport_type_6[ax].widthloss_code
    

      ws.getCell("O"+count6).value = Number(this.rowexport_type_6[ax].endless_length_yd)
      ws.getCell("O"+count6).numFmt = '0.00';

      ws.getCell("P"+count6).value = this.rowexport_type_6[ax].avg_end/100
      ws.getCell("P"+count6).numFmt = '0.00%';

      ws.getCell("Q"+count6).value = this.rowexport_type_6[ax].avg_end_code
      

      ws.getCell("R"+count6).value = Number(this.rowexport_type_6[ax].spliceloss)
      ws.getCell("R"+count6).numFmt = '0.00';

      ws.getCell("S"+count6).value = this.rowexport_type_6[ax].splicelossper/100
      ws.getCell("S"+count6).numFmt = '0.00%';

      ws.getCell("T"+count6).value = this.rowexport_type_6[ax].per_splice_loss_code
   

      ws.getCell("U"+count6).value = Number(this.rowexport_type_6[ax].totalcutout)
      ws.getCell("U"+count6).numFmt = '0.00';

      ws.getCell("V"+count6).value = this.rowexport_type_6[ax].totalcutoutper/100
      ws.getCell("V"+count6).numFmt = '0.00%';

      ws.getCell("W"+count6).value = this.rowexport_type_6[ax].total_cut_out_per_code

      ws.getCell("X"+count6).value = Number(this.rowexport_type_6[ax].rement)
      ws.getCell("X"+count6).numFmt = '0.00';

      ws.getCell("Y"+count6).value = this.rowexport_type_6[ax].rement_per/100
      ws.getCell("Y"+count6).numFmt = '0.00%';

      ws.getCell("Z"+count6).value = this.rowexport_type_6[ax].remnants_loss_code

      ws.getCell("AA"+count6).value = Number(this.rowexport_type_6[ax].totallenthloss)
      ws.getCell("AA"+count6).numFmt = '0.00';

      ws.getCell("AB"+count6).value = this.rowexport_type_6[ax].totallenthlossper/100
      ws.getCell("AB"+count6).numFmt = '0.00%';
      
      this.sum_plug12_t6 = 0
      this.sum_ttl_marker_t6 = Number(this.sum_ttl_marker_t6) + Number(this.rowexport_type_6[ax].ttl_marker)
      this.sum_plug12_t6 = this.rowexport_type_6[ax].plug12 * this.rowexport_type_6[ax].ttl_marker
      this.sum_per_width_loss_t6 = this.rowexport_type_6[ax].widthloss * this.rowexport_type_6[ax].ttl_marker
      this.sum_end1_2_t6 = this.rowexport_type_6[ax].end1_2 * this.rowexport_type_6[ax].ttl_marker
      this.sum_endless_length_yd_t6 = Number(this.sum_endless_length_yd_t6) + Number(this.rowexport_type_6[ax].endless_length_yd)
      this.sum_spliceloss_t6 = Number(this.sum_spliceloss_t6) + Number(this.rowexport_type_6[ax].spliceloss)
      this.sum_totalcutout_t6 = Number(this.sum_totalcutout_t6) + Number(this.rowexport_type_6[ax].totalcutout)
      this.sum_rement_t6 = Number(this.sum_rement_t6) + Number(this.rowexport_type_6[ax].rement)
      this.summary_plug12_t6 = Number(this.summary_plug12_t6) + Number(this.sum_plug12_t6)
      this.summary_widthloss_t6 = Number(this.summary_widthloss_t6) + Number(this.sum_per_width_loss_t6)
      this.summary_end1_2_t6 = Number(this.summary_end1_2_t6) + Number(this.sum_end1_2_t6)
      this.sum_totallenthloss_t6_first = 0
      this.sum_totallenthloss_t6_first = Number(this.rowexport_type_6[ax].endless_length_yd) + Number(this.rowexport_type_6[ax].spliceloss)
      +Number(this.rowexport_type_6[ax].totalcutout) + Number(this.rowexport_type_6[ax].rement)
      this.sum_totallenthloss_t6 = Number(this.sum_totallenthloss_t6) + Number(this.sum_totallenthloss_t6_first)

      this.sum_totallenthloss_per_t6_first = 0
      this.sum_totallenthloss_per_t6_first = Number(this.rowexport_type_6[ax].avg_end) + Number(this.rowexport_type_6[ax].splicelossper)
      +Number(this.rowexport_type_6[ax].totalcutoutper) + Number(this.rowexport_type_6[ax].rement_per)
      this.sum_totallenthloss_per_t6 = Number(this.sum_totallenthloss_t6) + Number(this.sum_totallenthloss_per_t6_first)
      this.sum_number_1_end_t6 = 0 
      this.sum_number_1_end_t6 = this.rowexport_type_6[ax].endless_length_yd * 36
      this.sum_number_2_end_t6 = 0
      this.sum_number_2_end_t6 = this.rowexport_type_6[ax].ttl_marker/this.rowexport_type_6[ax].length_ydb
      this.sum_plug1_2_end_t6 = 0
      this.sum_plug1_2_end_t6 = this.sum_number_1_end_t6/this.sum_number_2_end_t6

      this.sum_plug1_2_end_total_t6 = 0
     
      this.sum_plug1_2_end_total_t6 = this.sum_plug1_2_end_t6 * this.rowexport_type_6[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t6 = Number(this.sum_all_plug1_2_end_total_t6) + Number(this.sum_plug1_2_end_total_t6)

      this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t6
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t6
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t6
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t6
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t6
      })

       this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t6,
        value_table:6
      })
      ws.getCell("N"+count6).value = Number(this.sum_plug1_2_end_t6)
      ws.getCell("N"+count6).numFmt = '0.00';

      count6++
         }
         if(this.sum_ttl_marker_t6 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t6
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t6 = this.summary_plug12_t6 / this.sum_ttl_marker_t6
        if(this.sum_plug12_t6 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t6
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t6 =  this.summary_widthloss_t6 / this.sum_ttl_marker_t6
        if(this.sum_per_width_loss_t6 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t6
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t6 =  this.summary_end1_2_t6/ this.sum_ttl_marker_t6 //M
      if(this.sum_end1_2_t6 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t6
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t6 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t6
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t6 = this.sum_endless_length_yd_t6 / this.sum_ttl_marker_t6 *100 //P
      if(this.sum_avg_end_t6 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t6
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t6 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t6
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t6 = this.sum_spliceloss_t6 / this.sum_ttl_marker_t6 *100
      if(this.sum_splicelossper_t6 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t6
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t6 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t6
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t6 = this.sum_totalcutout_t6 / this.sum_ttl_marker_t6 *100
      if(this.sum_totalcutoutper_t6 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t6
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t6 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t6
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t6 = this.sum_rement_t6 / this.sum_ttl_marker_t6 *100
  if(this.sum_rement_per_t6 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t6
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t6
    if(this.sum_totallenthloss_t6 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t6
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert6 = this.sum_totallenthlosst6 / this.sum_ttl_marker_t6
    if(this.sum_totallenthlosspert6 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert6
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }


       ws.getCell("A"+countx6).value = "Total Table A6"  
      ws.mergeCells("A"+countx6+":"+"I"+countx6)
      
      if(this.sum_ttl_marker_t6 > 0 && isNaN(this.sum_ttl_marker_t6)==false && isFinite(this.sum_ttl_marker_t6)==true){
      ws.getCell("J"+countx6).value = Number(this.sum_ttl_marker_t6)
      ws.getCell("J"+countx6).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx6).value = 0
      ws.getCell("J"+countx6).numFmt = '0.00';
      }

      if(this.sum_plug12_t6 > 0 && isNaN(this.sum_plug12_t6)==false && isFinite(this.sum_plug12_t6)==true){
      ws.getCell("K"+countx6).value = Number(this.sum_plug12_t6)
      ws.getCell("K"+countx6).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx6).value = 0
      ws.getCell("K"+countx6).numFmt = '0.00';
      }

      if(this.sum_per_width_loss_t6 > 0 && isNaN(this.sum_per_width_loss_t6)==false && isFinite(this.sum_per_width_loss_t6)==true){
      ws.getCell("L"+countx6).value = this.sum_per_width_loss_t6/100
      ws.getCell("L"+countx6).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx6).value = 0.00/100
      ws.getCell("L"+countx6).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t6 = 0
      this.total_result_plug1_2_t6 = this.sum_all_plug1_2_end_total_t6 / this.sum_ttl_marker_t6

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t6
      })
      if(this.total_result_plug1_2_t6 > 0 && isNaN(this.total_result_plug1_2_t6)==false && isFinite(this.total_result_plug1_2_t6)==true){
      ws.getCell("N"+countx6).value = Number(this.total_result_plug1_2_t6)
      ws.getCell("N"+countx6).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx6).value = 0.00
      ws.getCell("N"+countx6).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_t6 > 0 && isNaN(this.sum_endless_length_yd_t6)==false && isFinite(this.sum_endless_length_yd_t6)==true){
      ws.getCell("O"+countx6).value = Number(this.sum_endless_length_yd_t6)
      ws.getCell("O"+countx6).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx6).value = 0.00
      ws.getCell("O"+countx6).numFmt = '0.00';
      }

      if(this.sum_avg_end_t6 > 0 && isNaN(this.sum_avg_end_t6)==false && isFinite(this.sum_avg_end_t6)==true){
      ws.getCell("P"+countx6).value = this.sum_avg_end_t6/100
      ws.getCell("P"+countx6).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx6).value = 0/100
      ws.getCell("P"+countx6).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t6 > 0 && isNaN(this.sum_spliceloss_t6)==false && isFinite(this.sum_spliceloss_t6)==true){
      ws.getCell("R"+countx6).value = Number(this.sum_spliceloss_t6)
      ws.getCell("R"+countx6).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx6).value = 0.00
      ws.getCell("R"+countx6).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t6 > 0 && isNaN(this.sum_splicelossper_t6)==false && isFinite(this.sum_splicelossper_t6)==true){
      ws.getCell("S"+countx6).value = this.sum_splicelossper_t6/100
      ws.getCell("S"+countx6).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx6).value = 0/100
      ws.getCell("S"+countx6).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t6 > 0 && isNaN(this.sum_totalcutout_t6)==false && isFinite(this.sum_totalcutout_t6)==true){
      ws.getCell("U"+countx6).value = Number(this.sum_totalcutout_t6)
      ws.getCell("U"+countx6).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx6).value = 0.00
      ws.getCell("U"+countx6).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t6 > 0 && isNaN(this.sum_totalcutoutper_t6)==false && isFinite(this.sum_totalcutoutper_t6)==true){
      ws.getCell("V"+countx6).value = this.sum_totalcutoutper_t6/100
      ws.getCell("V"+countx6).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx6).value = 0.00/100
      ws.getCell("V"+countx6).numFmt = '0.00%';
      }

       if(this.sum_rement_t6 > 0 && isNaN(this.sum_rement_t6)==false && isFinite(this.sum_rement_t6)==true){
      ws.getCell("X"+countx6).value = Number(this.sum_rement_t6)
      ws.getCell("X"+countx6).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx6).value = 0.00
      ws.getCell("X"+countx6).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t6 > 0 && isNaN(this.sum_rement_per_t6)==false && isFinite(this.sum_rement_per_t6)==true){
      ws.getCell("Y"+countx6).value =  this.sum_rement_per_t6/100
      ws.getCell("Y"+countx6).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx6).value =  0.00/100
      ws.getCell("Y"+countx6).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t6 > 0 && isNaN(this.sum_totallenthloss_t6)==false && isFinite(this.sum_totallenthloss_t6)==true){
      ws.getCell("AA"+countx6).value = Number(this.sum_totallenthloss_t6)
      ws.getCell("AA"+countx6).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx6).value = 0
      ws.getCell("AA"+countx6).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t6 = this.sum_totallenthloss_t6 / this.sum_ttl_marker_t6 *100
      if(this.last_sum_totallenthloss_per_t6 > 0 && isNaN(this.last_sum_totallenthloss_per_t6)==false && isFinite(this.last_sum_totallenthloss_per_t6)==true){
      ws.getCell("AB"+countx6).value = this.last_sum_totallenthloss_per_t6/100
      ws.getCell("AB"+countx6).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx6).value = 0/100
      ws.getCell("AB"+countx6).numFmt = '0.00%'; 
      }
      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx6).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //A7
    
  if(this.rowexport_type_6.length == 0){
       
         var count7 = count6

       }else{
      var count7 = countx6+1
       }
        var countx7 = count7 + 1 +Number(this.rowexport_type_7.length)

        
      
      this.sum_ttl_marker_t7 = 0
      this.sum_plug12_t7 = 0
      this.sum_per_width_loss_t7 = 0
      this.sum_end1_2_t7 = 0
      this.sum_endless_length_yd_t7= 0
      this.sum_avg_end_t7 = 0
      this.sum_spliceloss_t7 = 0
      this.sum_splicelossper_t7 = 0
      this.sum_totalcutout_t7 = 0
      this.sum_totalcutoutper_t7 = 0
      this.sum_rement_t7 = 0
      this.sum_rement_per_t7 = 0
      this.sum_totallenthloss_t7 =0 
      this.sum_totallenthlossper_t7 = 0
      this.summary_widthloss_t7 = 0
      this.summary_plug12_t7 = 0
      this.summary_end1_2_t7 = 0

      
     
      this.sum_all_plug1_2_end_total_t7 = 0

      if(this.rowexport_type_7.length < 1){
        count_row_grand++
      }else{
          for (var ax = 0; ax<this.rowexport_type_7.length; ax++){ //length +70
      ws.getCell("A"+count7).value = this.rowexport_type_7[ax].mu_date

      ws.getCell("B"+count7).value = this.rowexport_type_7[ax].gm_table

      ws.getCell("C"+count7).value = Number(this.rowexport_type_7[ax].gm_no_short)
      ws.getCell("C"+count7).numFmt = '0.00';

      ws.getCell("D"+count7).value = this.rowexport_type_7[ax].customer

      ws.getCell("E"+count7).value = this.rowexport_type_7[ax].gm_no

      ws.getCell("F"+count7).value = this.rowexport_type_7[ax].fabric_type

      ws.getCell("G"+count7).value = Number(this.rowexport_type_7[ax].obs_width)
      ws.getCell("G"+count7).numFmt = '0.00';

      ws.getCell("H"+count7).value = Number(this.rowexport_type_7[ax].width_inch)
      ws.getCell("H"+count7).numFmt = '0.00';


      ws.getCell("I"+count7).value = Number(this.rowexport_type_7[ax].length_ydb)
      ws.getCell("I"+count7).numFmt = '0.00';

      ws.getCell("J"+count7).value = Number(this.rowexport_type_7[ax].ttl_marker)
      ws.getCell("J"+count7).numFmt = '0.00';

      ws.getCell("K"+count7).value = Number(this.rowexport_type_7[ax].plug12)
      ws.getCell("K"+count7).numFmt = '0.00';

      ws.getCell("L"+count7).value = this.rowexport_type_7[ax].widthloss/100
      ws.getCell("L"+count7).numFmt = '0.00%';

      ws.getCell("M"+count7).value = this.rowexport_type_7[ax].widthloss_code

      ws.getCell("O"+count7).value = Number(this.rowexport_type_7[ax].endless_length_yd)
      ws.getCell("O"+count7).numFmt = '0.00';

      ws.getCell("P"+count7).value = this.rowexport_type_7[ax].avg_end/100
      ws.getCell("P"+count7).numFmt = '0.00%';

      ws.getCell("Q"+count7).value = this.rowexport_type_7[ax].avg_end_code
      

      ws.getCell("R"+count7).value = Number(this.rowexport_type_7[ax].spliceloss)
      ws.getCell("R"+count7).numFmt = '0.00';

      ws.getCell("S"+count7).value = this.rowexport_type_7[ax].splicelossper/100
      ws.getCell("S"+count7).numFmt = '0.00%';

      ws.getCell("T"+count7).value = this.rowexport_type_7[ax].per_splice_loss_code
   

      ws.getCell("U"+count7).value = Number(this.rowexport_type_7[ax].totalcutout)
      ws.getCell("U"+count7).numFmt = '0.00';

      ws.getCell("V"+count7).value = this.rowexport_type_7[ax].totalcutoutper/100
      ws.getCell("V"+count7).numFmt = '0.00%';

      ws.getCell("W"+count7).value = this.rowexport_type_7[ax].total_cut_out_per_code

      ws.getCell("X"+count7).value = Number(this.rowexport_type_7[ax].rement)
      ws.getCell("X"+count7).numFmt = '0.00';

      ws.getCell("Y"+count7).value = this.rowexport_type_7[ax].rement_per/100
      ws.getCell("Y"+count7).numFmt = '0.00%';

      ws.getCell("Z"+count7).value = this.rowexport_type_7[ax].remnants_loss_code

      ws.getCell("AA"+count7).value = Number(this.rowexport_type_7[ax].totallenthloss)
      ws.getCell("AA"+count7).numFmt = '0.00';

      ws.getCell("AB"+count7).value = this.rowexport_type_7[ax].totallenthlossper/100
      ws.getCell("AB"+count7).numFmt = '0.00%';
      
      
      this.sum_plug12_t7 = 0
      this.sum_ttl_marker_t7 = Number(this.sum_ttl_marker_t7) + Number(this.rowexport_type_7[ax].ttl_marker)
      this.sum_plug12_t7 = this.rowexport_type_7[ax].plug12 * this.rowexport_type_7[ax].ttl_marker
      this.sum_per_width_loss_t7 = this.rowexport_type_7[ax].widthloss * this.rowexport_type_7[ax].ttl_marker
      this.sum_end1_2_t7 = this.rowexport_type_7[ax].end1_2 * this.rowexport_type_7[ax].ttl_marker
      this.sum_endless_length_yd_t7 = Number(this.sum_endless_length_yd_t7) + Number(this.rowexport_type_7[ax].endless_length_yd)
      this.sum_spliceloss_t7 = Number(this.sum_spliceloss_t7) + Number(this.rowexport_type_7[ax].spliceloss)
      this.sum_totalcutout_t7 = Number(this.sum_totalcutout_t7) + Number(this.rowexport_type_7[ax].totalcutout)
      this.sum_rement_t7 = Number(this.sum_rement_t7) + Number(this.rowexport_type_7[ax].rement)
      this.summary_plug12_t7 = Number(this.summary_plug12_t7) + Number(this.sum_plug12_t7)
      this.summary_widthloss_t7 = Number(this.summary_widthloss_t7) + Number(this.sum_per_width_loss_t7)
      this.summary_end1_2_t7 = Number(this.summary_end1_2_t7) + Number(this.sum_end1_2_t7)
      this.sum_totallenthloss_t7_first = 0
      this.sum_totallenthloss_t7_first = Number(this.rowexport_type_7[ax].endless_length_yd) + Number(this.rowexport_type_7[ax].spliceloss)
      +Number(this.rowexport_type_7[ax].totalcutout) + Number(this.rowexport_type_7[ax].rement)
      this.sum_totallenthloss_t7 = Number(this.sum_totallenthloss_t7) + Number(this.sum_totallenthloss_t7_first)

      this.sum_totallenthloss_per_t7_first = 0
      this.sum_totallenthloss_per_t7_first = Number(this.rowexport_type_7[ax].avg_end) + Number(this.rowexport_type_7[ax].splicelossper)
      +Number(this.rowexport_type_7[ax].totalcutoutper) + Number(this.rowexport_type_7[ax].rement_per)
      this.sum_totallenthloss_per_t7 = Number(this.sum_totallenthloss_t7) + Number(this.sum_totallenthloss_per_t7_first)
      this.sum_number_1_end_t7 = 0 
      this.sum_number_1_end_t7 = this.rowexport_type_7[ax].endless_length_yd * 36
      this.sum_number_2_end_t7 = 0
      this.sum_number_2_end_t7 = this.rowexport_type_7[ax].ttl_marker/this.rowexport_type_7[ax].length_ydb
      this.sum_plug1_2_end_t7 = 0
      this.sum_plug1_2_end_t7 = this.sum_number_1_end_t7/this.sum_number_2_end_t7
      
      this.sum_plug1_2_end_total_t7 = 0
     
      this.sum_plug1_2_end_total_t7 = this.sum_plug1_2_end_t7 * this.rowexport_type_7[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t7 = Number(this.sum_all_plug1_2_end_total_t7) + Number(this.sum_plug1_2_end_total_t7)

      this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t7
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t7
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t7
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t7
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t7
      })

      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t7,
        value_table:7
      })
       
      ws.getCell("N"+count7).value = Number(this.sum_plug1_2_end_t7)
      ws.getCell("N"+count7).numFmt = '0.00';

      count7++
         }
         if(this.sum_ttl_marker_t7 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t7
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t7 = this.summary_plug12_t7 / this.sum_ttl_marker_t7
        if(this.sum_plug12_t7 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t7
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t7 =  this.summary_widthloss_t7 / this.sum_ttl_marker_t7
        if(this.sum_per_width_loss_t7 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t7
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t7 =  this.sum_end1_2_t7/ this.sum_ttl_marker_t7 //M
      if(this.sum_end1_2_t7 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t7
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t7 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t7
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t7 = this.sum_endless_length_yd_t7 / this.sum_ttl_marker_t7 *100 //P
      if(this.sum_avg_end_t7 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t7
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t7 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t7
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t7 = this.sum_spliceloss_t7 / this.sum_ttl_marker_t7 *100
      if(this.sum_splicelossper_t7 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t7
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t7 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t7
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t7 = this.sum_totalcutout_t7 / this.sum_ttl_marker_t7 *100
      if(this.sum_totalcutoutper_t7 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t7
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t7 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t7
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t7 = this.sum_rement_t7 / this.sum_ttl_marker_t7 *100
  if(this.sum_rement_per_t7 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t7
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t7
    if(this.sum_totallenthloss_t7 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t7
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert7 = this.sum_totallenthlosst7 / this.sum_ttl_marker_t7
    if(this.sum_totallenthlosspert7 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert7
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }


       ws.getCell("A"+countx7).value = "Total Table A7"  
      ws.mergeCells("A"+countx7+":"+"I"+countx7)
      
      if(this.sum_ttl_marker_t7 > 0 && isNaN(this.sum_ttl_marker_t7)==false && isFinite(this.sum_ttl_marker_t7)==true){
      ws.getCell("J"+countx7).value = Number(this.sum_ttl_marker_t7)
      ws.getCell("J"+countx7).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx7).value = 0
      ws.getCell("J"+countx7).numFmt = '0.00';
      }

      if(this.sum_plug12_t7 > 0 && isNaN(this.sum_plug12_t7)==false && isFinite(this.sum_plug12_t7)==true){
      ws.getCell("K"+countx7).value = Number(this.sum_plug12_t7)
      ws.getCell("K"+countx7).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx7).value = 0
      ws.getCell("K"+countx7).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t7 > 0 && isNaN(this.sum_per_width_loss_t7)==false && isFinite(this.sum_per_width_loss_t7)==true){
      ws.getCell("L"+countx7).value = this.sum_per_width_loss_t7/100
      ws.getCell("L"+countx7).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx7).value = 0.00/100
      ws.getCell("L"+countx7).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t7 = 0
      this.total_result_plug1_2_t7 = this.sum_all_plug1_2_end_total_t7 / this.sum_ttl_marker_t7

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t7
      })
      if(this.total_result_plug1_2_t7 > 0 && isNaN(this.total_result_plug1_2_t7)==false && isFinite(this.total_result_plug1_2_t7)==true){
      ws.getCell("N"+countx7).value = Number(this.total_result_plug1_2_t7)
      ws.getCell("N"+countx7).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx7).value = 0.00
      ws.getCell("N"+countx7).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_t7 > 0 && isNaN(this.sum_endless_length_yd_t7)==false && isFinite(this.sum_endless_length_yd_t7)==true){
      ws.getCell("O"+countx7).value = Number(this.sum_endless_length_yd_t7)
      ws.getCell("O"+countx7).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx7).value = 0.00
      ws.getCell("O"+countx7).numFmt = '0.00';
      }

      if(this.sum_avg_end_t7 > 0 && isNaN(this.sum_avg_end_t7)==false && isFinite(this.sum_avg_end_t7)==true){
      ws.getCell("P"+countx7).value = this.sum_avg_end_t7/100
      ws.getCell("P"+countx7).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx7).value = 0/100
      ws.getCell("P"+countx7).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t7 > 0 && isNaN(this.sum_spliceloss_t7)==false && isFinite(this.sum_spliceloss_t7)==true){
      ws.getCell("R"+countx7).value = Number(this.sum_spliceloss_t7)
      ws.getCell("R"+countx7).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx7).value = 0.00
      ws.getCell("R"+countx7).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t7 > 0 && isNaN(this.sum_splicelossper_t7)==false && isFinite(this.sum_splicelossper_t7)==true){
      ws.getCell("S"+countx7).value = this.sum_splicelossper_t7/100
      ws.getCell("S"+countx7).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx7).value = 0/100
      ws.getCell("S"+countx7).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t7 > 0 && isNaN(this.sum_totalcutout_t7)==false && isFinite(this.sum_totalcutout_t7)==true){
      ws.getCell("U"+countx7).value = Number(this.sum_totalcutout_t7)
      ws.getCell("U"+countx7).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx7).value = 0.00
      ws.getCell("U"+countx7).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t7 > 0 && isNaN(this.sum_totalcutoutper_t7)==false && isFinite(this.sum_totalcutoutper_t7)==true){
      ws.getCell("V"+countx7).value = this.sum_totalcutoutper_t7/100
      ws.getCell("V"+countx7).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx7).value = 0.00/100
      ws.getCell("V"+countx7).numFmt = '0.00%';
      }

       if(this.sum_rement_t7 > 0 && isNaN(this.sum_rement_t7)==false && isFinite(this.sum_rement_t7)==true){
      ws.getCell("X"+countx7).value = Number(this.sum_rement_t7)
      ws.getCell("X"+countx7).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx7).value = 0.00
      ws.getCell("X"+countx7).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t7 > 0 && isNaN(this.sum_rement_per_t7)==false && isFinite(this.sum_rement_per_t7)==true){
      ws.getCell("Y"+countx7).value =  this.sum_rement_per_t7/100
      ws.getCell("Y"+countx7).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx7).value =  0.00/100
      ws.getCell("Y"+countx7).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t7 > 0 && isNaN(this.sum_totallenthloss_t7)==false && isFinite(this.sum_totallenthloss_t7)==true){
      ws.getCell("AA"+countx7).value = Number(this.sum_totallenthloss_t7)
      ws.getCell("AA"+countx7).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx7).value = 0
      ws.getCell("AA"+countx7).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t7 = this.sum_totallenthloss_t7 / this.sum_ttl_marker_t7 *100
      if(this.last_sum_totallenthloss_per_t7 > 0 && isNaN(this.last_sum_totallenthloss_per_t7)==false && isFinite(this.last_sum_totallenthloss_per_t7)==true){
      ws.getCell("AB"+countx7).value = this.last_sum_totallenthloss_per_t7/100
      ws.getCell("AB"+countx7).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx7).value = 0/100
      ws.getCell("AB"+countx7).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx7).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //A8
     
    if(this.rowexport_type_7.length == 0){
       
         var count8 = count7

       }else{
      var count8 = countx7+1
       }
        var countx8 = count8 + 1 +Number(this.rowexport_type_8.length)
    
      this.sum_ttl_marker_t8 = 0
      this.sum_plug12_t8 = 0
      this.sum_per_width_loss_t8 = 0
      this.sum_end1_2_t8 = 0
      this.sum_endless_length_yd_t8= 0
      this.sum_avg_end_t8 = 0
      this.sum_spliceloss_t8 = 0
      this.sum_splicelossper_t8 = 0
      this.sum_totalcutout_t8 = 0
      this.sum_totalcutoutper_t8 = 0
      this.sum_rement_t8 = 0
      this.sum_rement_per_t8 = 0
      this.sum_totallenthloss_t8 =0 
      this.sum_totallenthlossper_t8 = 0
      this.summary_widthloss_t8 = 0
      this.summary_plug12_t8 = 0
      this.summary_end1_2_t8 = 0

      
   
     
      this.sum_all_plug1_2_end_total_t8 = 0

      if(this.rowexport_type_8.length < 1){
        count_row_grand++
      }else{
          for (var ax = 0; ax<this.rowexport_type_8.length; ax++){ //length +80
      ws.getCell("A"+count8).value = this.rowexport_type_8[ax].mu_date

      ws.getCell("B"+count8).value = this.rowexport_type_8[ax].gm_table

      ws.getCell("C"+count8).value = Number(this.rowexport_type_8[ax].gm_no_short)
      ws.getCell("C"+count8).numFmt = '0.00';

      ws.getCell("D"+count8).value = this.rowexport_type_8[ax].customer

      ws.getCell("E"+count8).value = this.rowexport_type_8[ax].gm_no

      ws.getCell("F"+count8).value = this.rowexport_type_8[ax].fabric_type

      ws.getCell("G"+count8).value = Number(this.rowexport_type_8[ax].obs_width)
      ws.getCell("G"+count8).numFmt = '0.00';

      ws.getCell("H"+count8).value = Number(this.rowexport_type_8[ax].width_inch)
      ws.getCell("H"+count8).numFmt = '0.00';


      ws.getCell("I"+count8).value = Number(this.rowexport_type_8[ax].length_ydb)
      ws.getCell("I"+count8).numFmt = '0.00';

      ws.getCell("J"+count8).value = Number(this.rowexport_type_8[ax].ttl_marker)
      ws.getCell("J"+count8).numFmt = '0.00';

      ws.getCell("K"+count8).value = Number(this.rowexport_type_8[ax].plug12)
      ws.getCell("K"+count8).numFmt = '0.00';

      
      ws.getCell("L"+count8).value = this.rowexport_type_8[ax].widthloss/100
      ws.getCell("L"+count8).numFmt = '0.00%';

      ws.getCell("M"+count8).value = this.rowexport_type_8[ax].widthloss_code

      ws.getCell("O"+count8).value = Number(this.rowexport_type_8[ax].endless_length_yd)
      ws.getCell("O"+count8).numFmt = '0.00';

      console.log(this.rowexport_type_8[ax].avg_end)
      ws.getCell("P"+count8).value = this.rowexport_type_8[ax].avg_end/100
      ws.getCell("P"+count8).numFmt = '0.00%';

      ws.getCell("Q"+count8).value = this.rowexport_type_8[ax].avg_end_code
      

      ws.getCell("R"+count8).value = Number(this.rowexport_type_8[ax].spliceloss)
      ws.getCell("R"+count8).numFmt = '0.00';

      ws.getCell("S"+count8).value = this.rowexport_type_8[ax].splicelossper/100
      ws.getCell("S"+count8).numFmt = '0.00%';

      ws.getCell("T"+count8).value = this.rowexport_type_8[ax].per_splice_loss_code
   

      ws.getCell("U"+count8).value = Number(this.rowexport_type_8[ax].totalcutout)
      ws.getCell("U"+count8).numFmt = '0.00';

      ws.getCell("V"+count8).value = this.rowexport_type_8[ax].totalcutoutper/100
      ws.getCell("V"+count8).numFmt = '0.00%';

      ws.getCell("W"+count8).value = this.rowexport_type_8[ax].total_cut_out_per_code

      ws.getCell("X"+count8).value = Number(this.rowexport_type_8[ax].rement)
      ws.getCell("X"+count8).numFmt = '0.00';

      ws.getCell("Y"+count8).value = this.rowexport_type_8[ax].rement_per/100
      ws.getCell("Y"+count8).numFmt = '0.00%';

      ws.getCell("Z"+count8).value = this.rowexport_type_8[ax].remnants_loss_code

      ws.getCell("AA"+count8).value = Number(this.rowexport_type_8[ax].totallenthloss)
      ws.getCell("AA"+count8).numFmt = '0.00';

      ws.getCell("AB"+count8).value = this.rowexport_type_8[ax].totallenthlossper/100
      ws.getCell("AB"+count8).numFmt = '0.00%';
      
      this.sum_plug12_t8 = 0
      this.sum_ttl_marker_t8 = Number(this.sum_ttl_marker_t8) + Number(this.rowexport_type_8[ax].ttl_marker)
      this.sum_plug12_t8 = this.rowexport_type_8[ax].plug12 * this.rowexport_type_8[ax].ttl_marker
      this.sum_per_width_loss_t8 = this.rowexport_type_8[ax].widthloss * this.rowexport_type_8[ax].ttl_marker
      this.sum_end1_2_t8 = this.rowexport_type_8[ax].end1_2 * this.rowexport_type_8[ax].ttl_marker
      this.sum_endless_length_yd_t8 = Number(this.sum_endless_length_yd_t8) + Number(this.rowexport_type_8[ax].endless_length_yd)
      this.sum_spliceloss_t8 = Number(this.sum_spliceloss_t8) + Number(this.rowexport_type_8[ax].spliceloss)
      this.sum_totalcutout_t8 = Number(this.sum_totalcutout_t8) + Number(this.rowexport_type_8[ax].totalcutout)
      this.sum_rement_t8 = Number(this.sum_rement_t8) + Number(this.rowexport_type_8[ax].rement)
      this.summary_plug12_t8 = Number(this.summary_plug12_t8) + Number(this.sum_plug12_t8)
      this.summary_widthloss_t8 = Number(this.summary_widthloss_t8) + Number(this.sum_per_width_loss_t8)
      this.summary_end1_2_t8 = Number(this.summary_end1_2_t8) + Number(this.sum_end1_2_t8)
      this.sum_totallenthloss_t8_first = 0
      this.sum_totallenthloss_t8_first = Number(this.rowexport_type_8[ax].endless_length_yd) + Number(this.rowexport_type_8[ax].spliceloss)
      +Number(this.rowexport_type_8[ax].totalcutout) + Number(this.rowexport_type_8[ax].rement)
      this.sum_totallenthloss_t8 = Number(this.sum_totallenthloss_t8) + Number(this.sum_totallenthloss_t8_first)

      this.sum_totallenthloss_per_t8_first = 0
      this.sum_totallenthloss_per_t8_first = Number(this.rowexport_type_8[ax].avg_end) + Number(this.rowexport_type_8[ax].splicelossper)
      +Number(this.rowexport_type_8[ax].totalcutoutper) + Number(this.rowexport_type_8[ax].rement_per)
      this.sum_totallenthloss_per_t8 = Number(this.sum_totallenthloss_t8) + Number(this.sum_totallenthloss_per_t8_first)
      this.sum_number_1_end_t8 = 0 
      this.sum_number_1_end_t8 = this.rowexport_type_8[ax].endless_length_yd * 36
      this.sum_number_2_end_t8 = 0
      this.sum_number_2_end_t8 = this.rowexport_type_8[ax].ttl_marker/this.rowexport_type_8[ax].length_ydb
      this.sum_plug1_2_end_t8 = 0
      this.sum_plug1_2_end_t8 = this.sum_number_1_end_t8/this.sum_number_2_end_t8

      this.sum_plug1_2_end_total_t8 = 0
     
      this.sum_plug1_2_end_total_t8 = this.sum_plug1_2_end_t8 * this.rowexport_type_8[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t8 = Number(this.sum_all_plug1_2_end_total_t8) + Number(this.sum_plug1_2_end_total_t8)

      this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t8
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t8
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t8
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t8
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t8
      })

      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t8,
        value_table:8
      })

       
      ws.getCell("N"+count8).value = Number(this.sum_plug1_2_end_t8)
      ws.getCell("N"+count8).numFmt = '0.00';

      count8++
         }
         if(this.sum_ttl_marker_t8 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t8
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t8 = this.summary_plug12_t8 / this.sum_ttl_marker_t8
        if(this.sum_plug12_t8 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t8
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t8 =  this.summary_widthloss_t8/ this.sum_ttl_marker_t8
        if(this.sum_per_width_loss_t8 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t8
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t8 =  this.summary_end1_2_t8/ this.sum_ttl_marker_t8 //M
      if(this.sum_end1_2_t8 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t8
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t8 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t8
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t8 = this.sum_endless_length_yd_t8 / this.sum_ttl_marker_t8 *100 //P
      if(this.sum_avg_end_t8 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t8
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t8 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t8
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t8 = this.sum_spliceloss_t8 / this.sum_ttl_marker_t8 *100
      if(this.sum_splicelossper_t8 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t8
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t8 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t8
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t8 = this.sum_totalcutout_t8 / this.sum_ttl_marker_t8 *100
      if(this.sum_totalcutoutper_t8 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t8
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t8 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t8
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t8 = this.sum_rement_t8 / this.sum_ttl_marker_t8 *100
  if(this.sum_rement_per_t8 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t8
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t8
    if(this.sum_totallenthloss_t8 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t8
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert8 = this.sum_totallenthlosst8 / this.sum_ttl_marker_t8
    if(this.sum_totallenthlosspert8 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert8
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

ws.getCell("A"+countx8).value = "Total Table A8"  
      ws.mergeCells("A"+countx8+":"+"I"+countx8)
      
      if(this.sum_ttl_marker_t8 > 0 && isNaN(this.sum_ttl_marker_t8)==false && isFinite(this.sum_ttl_marker_t8)==true){
      ws.getCell("J"+countx8).value = Number(this.sum_ttl_marker_t8)
      ws.getCell("J"+countx8).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx8).value = 0
      ws.getCell("J"+countx8).numFmt = '0.00';
      }

      if(this.sum_plug12_t8 > 0 && isNaN(this.sum_plug12_t8)==false && isFinite(this.sum_plug12_t8)==true){
      ws.getCell("K"+countx8).value = Number(this.sum_plug12_t8)
      ws.getCell("K"+countx8).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx8).value = 0
      ws.getCell("K"+countx8).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t8 > 0 && isNaN(this.sum_per_width_loss_t8)==false && isFinite(this.sum_per_width_loss_t8)==true){
      ws.getCell("L"+countx8).value = this.sum_per_width_loss_t8/100
      ws.getCell("L"+countx8).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx8).value = 0.00/100
      ws.getCell("L"+countx8).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t8 = 0
      this.total_result_plug1_2_t8 = this.sum_all_plug1_2_end_total_t8 / this.sum_ttl_marker_t8

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t8
      })
      if(this.total_result_plug1_2_t8 > 0 && isNaN(this.total_result_plug1_2_t8)==false && isFinite(this.total_result_plug1_2_t8)==true){
      ws.getCell("N"+countx8).value = Number(this.total_result_plug1_2_t8)
      ws.getCell("N"+countx8).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx8).value = 0.00
      ws.getCell("N"+countx8).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_t8 > 0 && isNaN(this.sum_endless_length_yd_t8)==false && isFinite(this.sum_endless_length_yd_t8)==true){
      ws.getCell("O"+countx8).value = Number(this.sum_endless_length_yd_t8)
      ws.getCell("O"+countx8).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx8).value = 0.00
      ws.getCell("O"+countx8).numFmt = '0.00';
      }

      if(this.sum_avg_end_t8 > 0 && isNaN(this.sum_avg_end_t8)==false && isFinite(this.sum_avg_end_t8)==true){
      ws.getCell("P"+countx8).value = this.sum_avg_end_t8/100
      ws.getCell("P"+countx8).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx8).value = 0/100
      ws.getCell("P"+countx8).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t8 > 0 && isNaN(this.sum_spliceloss_t8)==false && isFinite(this.sum_spliceloss_t8)==true){
      ws.getCell("R"+countx8).value = Number(this.sum_spliceloss_t8)
      ws.getCell("R"+countx8).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx8).value = 0.00
      ws.getCell("R"+countx8).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t8 > 0 && isNaN(this.sum_splicelossper_t8)==false && isFinite(this.sum_splicelossper_t8)==true){
      ws.getCell("S"+countx8).value = this.sum_splicelossper_t8/100
      ws.getCell("S"+countx8).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx8).value = 0/100
      ws.getCell("S"+countx8).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t8 > 0 && isNaN(this.sum_totalcutout_t8)==false && isFinite(this.sum_totalcutout_t8)==true){
      ws.getCell("U"+countx8).value = Number(this.sum_totalcutout_t8)
      ws.getCell("U"+countx8).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx8).value = 0.00
      ws.getCell("U"+countx8).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t8 > 0 && isNaN(this.sum_totalcutoutper_t8)==false && isFinite(this.sum_totalcutoutper_t8)==true){
      ws.getCell("V"+countx8).value = this.sum_totalcutoutper_t8/100
      ws.getCell("V"+countx8).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx8).value = 0.00/100
      ws.getCell("V"+countx8).numFmt = '0.00%';
      }

       if(this.sum_rement_t8 > 0 && isNaN(this.sum_rement_t8)==false && isFinite(this.sum_rement_t8)==true){
      ws.getCell("X"+countx8).value = Number(this.sum_rement_t8)
      ws.getCell("X"+countx8).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx8).value = 0.00
      ws.getCell("X"+countx8).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t8 > 0 && isNaN(this.sum_rement_per_t8)==false && isFinite(this.sum_rement_per_t8)==true){
      ws.getCell("Y"+countx8).value =  this.sum_rement_per_t8/100
      ws.getCell("Y"+countx8).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx8).value =  0.00/100
      ws.getCell("Y"+countx8).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t8 > 0 && isNaN(this.sum_totallenthloss_t8)==false && isFinite(this.sum_totallenthloss_t8)==true){
      ws.getCell("AA"+countx8).value = Number(this.sum_totallenthloss_t8)
      ws.getCell("AA"+countx8).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx8).value = 0
      ws.getCell("AA"+countx8).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t8 = this.sum_totallenthloss_t8 / this.sum_ttl_marker_t8 *100
      if(this.last_sum_totallenthloss_per_t8 > 0 && isNaN(this.last_sum_totallenthloss_per_t8)==false && isFinite(this.last_sum_totallenthloss_per_t8)==true){
      ws.getCell("AB"+countx8).value = this.last_sum_totallenthloss_per_t8/100
      ws.getCell("AB"+countx8).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx8).value = 0/100
      ws.getCell("AB"+countx8).numFmt = '0.00%'; 
      }


      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx8).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //A9
     

      if(this.rowexport_type_8.length == 0){
       
         var count9 = count8

       }else{
      var count9 = countx8+1
       }
 
        var countx9 = count9 + 1 +Number(this.rowexport_type_9.length)
      
      this.sum_ttl_marker_t9 = 0
      this.sum_plug12_t9 = 0
      this.sum_per_width_loss_t9 = 0
      this.sum_end1_2_t9 = 0
      this.sum_endless_length_yd_t9= 0
      this.sum_avg_end_t9 = 0
      this.sum_spliceloss_t9 = 0
      this.sum_splicelossper_t9 = 0
      this.sum_totalcutout_t9 = 0
      this.sum_totalcutoutper_t9 = 0
      this.sum_rement_t9 = 0
      this.sum_rement_per_t9 = 0
      this.sum_totallenthloss_t9 =0 
      this.sum_totallenthlossper_t9 = 0
      this.summary_widthloss_t9 = 0
      this.summary_plug12_t9 = 0
      this.summary_end1_2_t9 = 0

      
    
     if(this.rowexport_type_9.length < 1){
        count_row_grand++
      }else{
      this.sum_all_plug1_2_end_total_t9 = 0
          for (var ax = 0; ax<this.rowexport_type_9.length; ax++){ //length +90
      ws.getCell("A"+count9).value = this.rowexport_type_9[ax].mu_date

      ws.getCell("B"+count9).value = this.rowexport_type_9[ax].gm_table

      ws.getCell("C"+count9).value = Number(this.rowexport_type_9[ax].gm_no_short)
      ws.getCell("C"+count9).numFmt = '0.00';

      ws.getCell("D"+count9).value = this.rowexport_type_9[ax].customer

      ws.getCell("E"+count9).value = this.rowexport_type_9[ax].gm_no

      ws.getCell("F"+count9).value = this.rowexport_type_9[ax].fabric_type

      ws.getCell("G"+count9).value = Number(this.rowexport_type_9[ax].obs_width)
      ws.getCell("G"+count9).numFmt = '0.00';

      ws.getCell("H"+count9).value = Number(this.rowexport_type_9[ax].width_inch)
      ws.getCell("H"+count9).numFmt = '0.00';


      ws.getCell("I"+count9).value = Number(this.rowexport_type_9[ax].length_ydb)
      ws.getCell("I"+count9).numFmt = '0.00';

      ws.getCell("J"+count9).value = Number(this.rowexport_type_9[ax].ttl_marker)
      ws.getCell("J"+count9).numFmt = '0.00';

      ws.getCell("K"+count9).value = Number(this.rowexport_type_9[ax].plug12)
      ws.getCell("K"+count9).numFmt = '0.00';

      ws.getCell("L"+count9).value = this.rowexport_type_9[ax].widthloss/100
      ws.getCell("L"+count9).numFmt = '0.00%';

      ws.getCell("M"+count9).value = this.rowexport_type_9[ax].widthloss_code
    
      ws.getCell("O"+count9).value = Number(this.rowexport_type_9[ax].endless_length_yd)
      ws.getCell("O"+count9).numFmt = '0.00';

      ws.getCell("P"+count9).value = this.rowexport_type_9[ax].avg_end/100
      ws.getCell("P"+count9).numFmt = '0.00%';

      ws.getCell("Q"+count9).value = this.rowexport_type_9[ax].avg_end_code
      

      ws.getCell("R"+count9).value = Number(this.rowexport_type_9[ax].spliceloss)
      ws.getCell("R"+count9).numFmt = '0.00';

      ws.getCell("S"+count9).value = this.rowexport_type_9[ax].splicelossper/100
      ws.getCell("S"+count9).numFmt = '0.00%';

      ws.getCell("T"+count9).value = this.rowexport_type_9[ax].per_splice_loss_code
   

      ws.getCell("U"+count9).value = Number(this.rowexport_type_9[ax].totalcutout)
      ws.getCell("U"+count9).numFmt = '0.00';

      ws.getCell("V"+count9).value = this.rowexport_type_9[ax].totalcutoutper/100
      ws.getCell("V"+count9).numFmt = '0.00%';

      ws.getCell("W"+count9).value = this.rowexport_type_9[ax].total_cut_out_per_code

      ws.getCell("X"+count9).value = Number(this.rowexport_type_9[ax].rement)
      ws.getCell("X"+count9).numFmt = '0.00';

      ws.getCell("Y"+count9).value = this.rowexport_type_9[ax].rement_per/100
      ws.getCell("Y"+count9).numFmt = '0.00%';

      ws.getCell("Z"+count9).value = this.rowexport_type_9[ax].remnants_loss_code

      ws.getCell("AA"+count9).value = Number(this.rowexport_type_9[ax].totallenthloss)
      ws.getCell("AA"+count9).numFmt = '0.00';

      ws.getCell("AB"+count9).value = this.rowexport_type_9[ax].totallenthlossper/100
      ws.getCell("AB"+count9).numFmt = '0.00%';
      
      this.sum_plug12_t9 = 0
      this.sum_ttl_marker_t9 = Number(this.sum_ttl_marker_t9) + Number(this.rowexport_type_9[ax].ttl_marker)
      this.sum_plug12_t9 = this.rowexport_type_9[ax].plug12 * this.rowexport_type_9[ax].ttl_marker
      this.sum_per_width_loss_t9 = this.rowexport_type_9[ax].widthloss * this.rowexport_type_9[ax].ttl_marker
      this.sum_end1_2_t9 = this.rowexport_type_9[ax].end1_2 * this.rowexport_type_9[ax].ttl_marker
      this.sum_endless_length_yd_t9 = Number(this.sum_endless_length_yd_t9) + Number(this.rowexport_type_9[ax].endless_length_yd)
      this.sum_spliceloss_t9 = Number(this.sum_spliceloss_t9) + Number(this.rowexport_type_9[ax].spliceloss)
      this.sum_totalcutout_t9 = Number(this.sum_totalcutout_t9) + Number(this.rowexport_type_9[ax].totalcutout)
      this.sum_rement_t9 = Number(this.sum_rement_t9) + Number(this.rowexport_type_9[ax].rement)
      this.summary_plug12_t9 = Number(this.summary_plug12_t9) + Number(this.sum_plug12_t9)
      this.summary_widthloss_t9 = Number(this.summary_widthloss_t9) + Number(this.sum_per_width_loss_t9)
      this.summary_end1_2_t9 = Number(this.summary_end1_2_t9) + Number(this.sum_end1_2_t9)
      this.sum_totallenthloss_t9_first = 0
      this.sum_totallenthloss_t9_first = Number(this.rowexport_type_9[ax].endless_length_yd) + Number(this.rowexport_type_9[ax].spliceloss)
      +Number(this.rowexport_type_9[ax].totalcutout) + Number(this.rowexport_type_9[ax].rement)
      this.sum_totallenthloss_t9 = Number(this.sum_totallenthloss_t9) + Number(this.sum_totallenthloss_t9_first)

      this.sum_totallenthloss_per_t9_first = 0
      this.sum_totallenthloss_per_t9_first = Number(this.rowexport_type_9[ax].avg_end) + Number(this.rowexport_type_9[ax].splicelossper)
      +Number(this.rowexport_type_9[ax].totalcutoutper) + Number(this.rowexport_type_9[ax].rement_per)
      this.sum_totallenthloss_per_t9 = Number(this.sum_totallenthloss_t9) + Number(this.sum_totallenthloss_per_t9_first)
      this.sum_number_1_end_t9 = 0 
      this.sum_number_1_end_t9 = this.rowexport_type_9[ax].endless_length_yd * 36
      this.sum_number_2_end_t9 = 0
      this.sum_number_2_end_t9 = this.rowexport_type_9[ax].ttl_marker/this.rowexport_type_9[ax].length_ydb
      this.sum_plug1_2_end_t9 = 0
      this.sum_plug1_2_end_t9 = this.sum_number_1_end_t9/this.sum_number_2_end_t9

      this.sum_plug1_2_end_total_t9 = 0
     
      this.sum_plug1_2_end_total_t9 = this.sum_plug1_2_end_t9 * this.rowexport_type_9[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t9 = Number(this.sum_all_plug1_2_end_total_t9) + Number(this.sum_plug1_2_end_total_t9)

      this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t9
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t9
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t9
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t9
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t9
      })
       
      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t9,
        value_table:9
      })

      ws.getCell("N"+count9).value = Number(this.sum_plug1_2_end_t9)
      ws.getCell("N"+count9).numFmt = '0.00';
      


      count9++
         }
         if(this.sum_ttl_marker_t9 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t9
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t9 = this.summary_plug12_t9 / this.sum_ttl_marker_t9
        if(this.sum_plug12_t9 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t9
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t9 =  this.summary_widthloss_t9/ this.sum_ttl_marker_t9
        if(this.sum_per_width_loss_t9 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t9
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t9 =  this.summary_end1_2_t9/ this.sum_ttl_marker_t9 //M
      if(this.sum_end1_2_t9 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t9
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t9 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t9
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t9 = this.sum_endless_length_yd_t9 / this.sum_ttl_marker_t9 *100 //P
      if(this.sum_avg_end_t9 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t9
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t9 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t9
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t9 = this.sum_spliceloss_t9 / this.sum_ttl_marker_t9 *100
      if(this.sum_splicelossper_t9 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t9
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t9 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t9
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t9 = this.sum_totalcutout_t9 / this.sum_ttl_marker_t9 *100
      if(this.sum_totalcutoutper_t9 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t9
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t9 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t9
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t9 = this.sum_rement_t9 / this.sum_ttl_marker_t9 *100
  if(this.sum_rement_per_t9 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t9
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t9
    if(this.sum_totallenthloss_t9 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t9
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert9 = this.sum_totallenthlosst9 / this.sum_ttl_marker_t9
    if(this.sum_totallenthlosspert9 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert9
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }


      ws.getCell("A"+countx9).value = "Total Table A9"  
      ws.mergeCells("A"+countx9+":"+"I"+countx9)
      
      if(this.sum_ttl_marker_t9 > 0 && isNaN(this.sum_ttl_marker_t9)==false && isFinite(this.sum_ttl_marker_t9)==true){
      ws.getCell("J"+countx9).value = Number(this.sum_ttl_marker_t9)
      ws.getCell("J"+countx9).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx9).value = 0
      ws.getCell("J"+countx9).numFmt = '0.00';
      }

      if(this.sum_plug12_t9 > 0 && isNaN(this.sum_plug12_t9)==false && isFinite(this.sum_plug12_t9)==true){
      ws.getCell("K"+countx9).value = Number(this.sum_plug12_t9)
      ws.getCell("K"+countx9).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx9).value = 0
      ws.getCell("K"+countx9).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t9 > 0 && isNaN(this.sum_per_width_loss_t9)==false && isFinite(this.sum_per_width_loss_t9)==true){
      ws.getCell("L"+countx9).value = this.sum_per_width_loss_t9/100
      ws.getCell("L"+countx9).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx9).value = 0.00/100
      ws.getCell("L"+countx9).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t9 = 0
      this.total_result_plug1_2_t9 = this.sum_all_plug1_2_end_total_t9 / this.sum_ttl_marker_t9

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t9
      })
     
      if(this.total_result_plug1_2_t9 > 0 && isNaN(this.total_result_plug1_2_t9)==false && isFinite(this.total_result_plug1_2_t9)==true){
      ws.getCell("N"+countx9).value = Number(this.total_result_plug1_2_t9)
      ws.getCell("N"+countx9).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx9).value = 0.00
      ws.getCell("N"+countx9).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_t9 > 0 && isNaN(this.sum_endless_length_yd_t9)==false && isFinite(this.sum_endless_length_yd_t9)==true){
      ws.getCell("O"+countx9).value = Number(this.sum_endless_length_yd_t9)
      ws.getCell("O"+countx9).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx9).value = 0.00
      ws.getCell("O"+countx9).numFmt = '0.00';
      }

      if(this.sum_avg_end_t9 > 0 && isNaN(this.sum_avg_end_t9)==false && isFinite(this.sum_avg_end_t9)==true){
      ws.getCell("P"+countx9).value = this.sum_avg_end_t9/100
      ws.getCell("P"+countx9).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx9).value = 0/100
      ws.getCell("P"+countx9).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t9 > 0 && isNaN(this.sum_spliceloss_t9)==false && isFinite(this.sum_spliceloss_t9)==true){
      ws.getCell("R"+countx9).value = Number(this.sum_spliceloss_t9)
      ws.getCell("R"+countx9).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx9).value = 0.00
      ws.getCell("R"+countx9).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t9 > 0 && isNaN(this.sum_splicelossper_t9)==false && isFinite(this.sum_splicelossper_t9)==true){
      ws.getCell("S"+countx9).value = this.sum_splicelossper_t9/100
      ws.getCell("S"+countx9).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx9).value = 0/100
      ws.getCell("S"+countx9).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t9 > 0 && isNaN(this.sum_totalcutout_t9)==false && isFinite(this.sum_totalcutout_t9)==true){
      ws.getCell("U"+countx9).value = Number(this.sum_totalcutout_t9)
      ws.getCell("U"+countx9).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx9).value = 0.00
      ws.getCell("U"+countx9).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t9 > 0 && isNaN(this.sum_totalcutoutper_t9)==false && isFinite(this.sum_totalcutoutper_t9)==true){
      ws.getCell("V"+countx9).value = this.sum_totalcutoutper_t9/100
      ws.getCell("V"+countx9).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx9).value = 0.00/100
      ws.getCell("V"+countx9).numFmt = '0.00%';
      }

       if(this.sum_rement_t9 > 0 && isNaN(this.sum_rement_t9)==false && isFinite(this.sum_rement_t9)==true){
      ws.getCell("X"+countx9).value = Number(this.sum_rement_t9)
      ws.getCell("X"+countx9).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx9).value = 0.00
      ws.getCell("X"+countx9).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t9 > 0 && isNaN(this.sum_rement_per_t9)==false && isFinite(this.sum_rement_per_t9)==true){
      ws.getCell("Y"+countx9).value =  this.sum_rement_per_t9/100
      ws.getCell("Y"+countx9).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx9).value =  0.00/100
      ws.getCell("Y"+countx9).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t9 > 0 && isNaN(this.sum_totallenthloss_t9)==false && isFinite(this.sum_totallenthloss_t9)==true){
      ws.getCell("AA"+countx9).value = Number(this.sum_totallenthloss_t9)
      ws.getCell("AA"+countx9).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx9).value = 0
      ws.getCell("AA"+countx9).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t9 = this.sum_totallenthloss_t9 / this.sum_ttl_marker_t9 *100
      if(this.last_sum_totallenthloss_per_t9 > 0 && isNaN(this.last_sum_totallenthloss_per_t9)==false && isFinite(this.last_sum_totallenthloss_per_t9)==true){
      ws.getCell("AB"+countx9).value = this.last_sum_totallenthloss_per_t9/100
      ws.getCell("AB"+countx9).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx9).value = 0/100
      ws.getCell("AB"+countx9).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx9).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //A10
      
    if(this.rowexport_type_9.length == 0){
       
         var count10 = count9

       }else{
      var count10 = countx9+1
       }
        var countx10 = count10 + 1 +Number(this.rowexport_type_10.length)
      
      this.sum_ttl_marker_t10 = 0
      this.sum_plug12_t10 = 0
      this.sum_per_width_loss_t10 = 0
      this.sum_end1_2_t10 = 0
      this.sum_endless_length_yd_t10= 0
      this.sum_avg_end_t10 = 0
      this.sum_spliceloss_t10 = 0
      this.sum_splicelossper_t10 = 0
      this.sum_totalcutout_t10 = 0
      this.sum_totalcutoutper_t10 = 0
      this.sum_rement_t10 = 0
      this.sum_rement_per_t10 = 0
      this.sum_totallenthloss_t10 =0 
      this.sum_totallenthlossper_t10 = 0
      this.summary_widthloss_t10 = 0
      this.summary_plug12_t10 = 0
      this.summary_end1_2_t10 = 0

      
   
      this.sum_all_plug1_2_end_total_t10 = 0

    if(this.rowexport_type_10.length < 1){ 
        count_row_grand++
      }else{
          for (var ax = 0; ax<this.rowexport_type_10.length; ax++){ //length +100
      ws.getCell("A"+count10).value = this.rowexport_type_10[ax].mu_date

      ws.getCell("B"+count10).value = this.rowexport_type_10[ax].gm_table

      ws.getCell("C"+count10).value = Number(this.rowexport_type_10[ax].gm_no_short)
      ws.getCell("C"+count10).numFmt = '0.00';

      ws.getCell("D"+count10).value = this.rowexport_type_10[ax].customer

      ws.getCell("E"+count10).value = this.rowexport_type_10[ax].gm_no

      ws.getCell("F"+count10).value = this.rowexport_type_10[ax].fabric_type

      ws.getCell("G"+count10).value = Number(this.rowexport_type_10[ax].obs_width)
      ws.getCell("G"+count10).numFmt = '0.00';

      ws.getCell("H"+count10).value = Number(this.rowexport_type_10[ax].width_inch)
      ws.getCell("H"+count10).numFmt = '0.00';


      ws.getCell("I"+count10).value = Number(this.rowexport_type_10[ax].length_ydb)
      ws.getCell("I"+count10).numFmt = '0.00';

      ws.getCell("J"+count10).value = Number(this.rowexport_type_10[ax].ttl_marker)
      ws.getCell("J"+count10).numFmt = '0.00';

      ws.getCell("K"+count10).value = Number(this.rowexport_type_10[ax].plug12)
      ws.getCell("K"+count10).numFmt = '0.00';

      ws.getCell("L"+count10).value = this.rowexport_type_10[ax].widthloss/100
      ws.getCell("L"+count10).numFmt = '0.00%';

      ws.getCell("M"+count10).value = this.rowexport_type_10[ax].widthloss_code
    

      ws.getCell("O"+count10).value = Number(this.rowexport_type_10[ax].endless_length_yd)
      ws.getCell("O"+count10).numFmt = '0.00';

      ws.getCell("P"+count10).value = this.rowexport_type_10[ax].avg_end/100
      ws.getCell("P"+count10).numFmt = '0.00%';

      ws.getCell("Q"+count10).value = this.rowexport_type_10[ax].avg_end_code
      

      ws.getCell("R"+count10).value = Number(this.rowexport_type_10[ax].spliceloss)
      ws.getCell("R"+count10).numFmt = '0.00';

      ws.getCell("S"+count10).value = this.rowexport_type_10[ax].splicelossper/100
      ws.getCell("S"+count10).numFmt = '0.00%';

      ws.getCell("T"+count10).value = this.rowexport_type_10[ax].per_splice_loss_code
   

      ws.getCell("U"+count10).value = Number(this.rowexport_type_10[ax].totalcutout)
      ws.getCell("U"+count10).numFmt = '0.00';

      ws.getCell("V"+count10).value = this.rowexport_type_10[ax].totalcutoutper/100
      ws.getCell("V"+count10).numFmt = '0.00%';

      ws.getCell("W"+count10).value = this.rowexport_type_10[ax].total_cut_out_per_code

      ws.getCell("X"+count10).value = Number(this.rowexport_type_10[ax].rement)
      ws.getCell("X"+count10).numFmt = '0.00';

      ws.getCell("Y"+count10).value = this.rowexport_type_10[ax].rement_per/100
      ws.getCell("Y"+count10).numFmt = '0.00%';

      ws.getCell("Z"+count10).value = this.rowexport_type_10[ax].remnants_loss_code

      ws.getCell("AA"+count10).value = Number(this.rowexport_type_10[ax].totallenthloss)
      ws.getCell("AA"+count10).numFmt = '0.00';

      ws.getCell("AB"+count10).value = this.rowexport_type_10[ax].totallenthlossper/100
      ws.getCell("AB"+count10).numFmt = '0.00%';
      
      this.sum_plug12_t10 = 0
      this.sum_ttl_marker_t10 = Number(this.sum_ttl_marker_t10) + Number(this.rowexport_type_10[ax].ttl_marker)
      this.sum_plug12_t10 = this.rowexport_type_10[ax].plug12 * this.rowexport_type_10[ax].ttl_marker

      this.sum_per_width_loss_t10 = this.rowexport_type_10[ax].widthloss * this.rowexport_type_10[ax].ttl_marker
      this.sum_end1_2_t10 = this.rowexport_type_10[ax].end1_2 * this.rowexport_type_10[ax].ttl_marker
      this.sum_endless_length_yd_t10 = Number(this.sum_endless_length_yd_t10) + Number(this.rowexport_type_10[ax].endless_length_yd)
      this.sum_spliceloss_t10 = Number(this.sum_spliceloss_t10) + Number(this.rowexport_type_10[ax].spliceloss)
      this.sum_totalcutout_t10 = Number(this.sum_totalcutout_t10) + Number(this.rowexport_type_10[ax].totalcutout)
      this.sum_rement_t10 = Number(this.sum_rement_t10) + Number(this.rowexport_type_10[ax].rement)
      this.summary_plug12_t10 = Number(this.summary_plug12_t10) + Number(this.sum_plug12_t10)
      this.summary_widthloss_t10 = Number(this.summary_widthloss_t10) + Number(this.sum_per_width_loss_t10)
      this.summary_end1_2_t10 = Number(this.summary_end1_2_t10) + Number(this.sum_end1_2_t10)
      this.sum_totallenthloss_t10_first = 0
      this.sum_totallenthloss_t10_first = Number(this.rowexport_type_10[ax].endless_length_yd) + Number(this.rowexport_type_10[ax].spliceloss)
      +Number(this.rowexport_type_10[ax].totalcutout) + Number(this.rowexport_type_10[ax].rement)
      this.sum_totallenthloss_t10 = Number(this.sum_totallenthloss_t10) + Number(this.sum_totallenthloss_t10_first)

      this.sum_totallenthloss_per_t10_first = 0
      this.sum_totallenthloss_per_t10_first = Number(this.rowexport_type_10[ax].avg_end) + Number(this.rowexport_type_10[ax].splicelossper)
      +Number(this.rowexport_type_10[ax].totalcutoutper) + Number(this.rowexport_type_10[ax].rement_per)
      this.sum_totallenthloss_per_t10 = Number(this.sum_totallenthloss_t10) + Number(this.sum_totallenthloss_per_t10_first)
      this.sum_number_1_end_t10 = 0 
      this.sum_number_1_end_t10 = this.rowexport_type_10[ax].endless_length_yd * 36
      this.sum_number_2_end_t10 = 0
      this.sum_number_2_end_t10 = this.rowexport_type_10[ax].ttl_marker/this.rowexport_type_10[ax].length_ydb
      this.sum_plug1_2_end_t10 = 0
      this.sum_plug1_2_end_t10 = this.sum_number_1_end_t10/this.sum_number_2_end_t10

      this.sum_plug1_2_end_total_t10 = 0
     
      this.sum_plug1_2_end_total_t10 = this.sum_plug1_2_end_t10 * this.rowexport_type_10[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t10 = Number(this.sum_all_plug1_2_end_total_t10) + Number(this.sum_plug1_2_end_total_t10)

      this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t10
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t10
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t10
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t10
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t10
      })

       this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t10,
        value_table:10
      })

      ws.getCell("N"+count10).value = Number(this.sum_plug1_2_end_t10)
      ws.getCell("N"+count10).numFmt = '0.00';

      count10++
         }
         if(this.sum_ttl_marker_t10 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t10
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t10 = this.summary_plug12_t10 / this.sum_ttl_marker_t10
        if(this.sum_plug12_t10 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t10
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t10 =  this.summary_widthloss_t10 / this.sum_ttl_marker_t10
        if(this.sum_per_width_loss_t10 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t10
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t10 =  this.summary_end1_2_t10/ this.sum_ttl_marker_t10 //M
      if(this.sum_end1_2_t10 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t10
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t10 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t10
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t10 = this.sum_endless_length_yd_t10 / this.sum_ttl_marker_t10 *100 //P
      if(this.sum_avg_end_t10 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t10
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t10 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t10
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t10 = this.sum_spliceloss_t10 / this.sum_ttl_marker_t10 *100
      if(this.sum_splicelossper_t10 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t10
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t10 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t10
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t10 = this.sum_totalcutout_t10 / this.sum_ttl_marker_t10 *100
      if(this.sum_totalcutoutper_t10 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t10
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t10 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t10
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t10 = this.sum_rement_t10 / this.sum_ttl_marker_t10 *100
  if(this.sum_rement_per_t10 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t10
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t10
    if(this.sum_totallenthloss_t10 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t10
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert10 = this.sum_totallenthlosst10 / this.sum_ttl_marker_t10
    if(this.sum_totallenthlosspert10 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert10
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

ws.getCell("A"+countx10).value = "Total Table A10"  
      ws.mergeCells("A"+countx10+":"+"I"+countx10)
      
      if(this.sum_ttl_marker_t10 > 0 && isNaN(this.sum_ttl_marker_t10)==false && isFinite(this.sum_ttl_marker_t10)==true){
      ws.getCell("J"+countx10).value = Number(this.sum_ttl_marker_t10)
      ws.getCell("J"+countx10).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx10).value = 0
      ws.getCell("J"+countx10).numFmt = '0.00';
      }

      if(this.sum_plug12_t10 > 0 && isNaN(this.sum_plug12_t10)==false && isFinite(this.sum_plug12_t10)==true){
      ws.getCell("K"+countx10).value = Number(this.sum_plug12_t10)
      ws.getCell("K"+countx10).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx10).value = 0
      ws.getCell("K"+countx10).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t10 > 0 && isNaN(this.sum_per_width_loss_t10)==false && isFinite(this.sum_per_width_loss_t10)==true){
      ws.getCell("L"+countx10).value = this.sum_per_width_loss_t10/100
      ws.getCell("L"+countx10).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx10).value = 0.00/100
      ws.getCell("L"+countx10).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t10 = 0
      this.total_result_plug1_2_t10 = this.sum_all_plug1_2_end_total_t10 / this.sum_ttl_marker_t10

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t10
      })
      if(this.total_result_plug1_2_t10 > 0 && isNaN(this.total_result_plug1_2_t10)==false && isFinite(this.total_result_plug1_2_t10)==true){
      ws.getCell("N"+countx10).value = Number(this.total_result_plug1_2_t10)
      ws.getCell("N"+countx10).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx10).value = 0.00
      ws.getCell("N"+countx10).numFmt = '0.00';
      }
      

      if(this.sum_endless_length_yd_t10 > 0 && isNaN(this.sum_endless_length_yd_t10)==false && isFinite(this.sum_endless_length_yd_t10)==true){
      ws.getCell("O"+countx10).value = Number(this.sum_endless_length_yd_t10)
      ws.getCell("O"+countx10).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx10).value = 0.00
      ws.getCell("O"+countx10).numFmt = '0.00';
      }

      if(this.sum_avg_end_t10 > 0 && isNaN(this.sum_avg_end_t10)==false && isFinite(this.sum_avg_end_t10)==true){
      ws.getCell("P"+countx10).value = this.sum_avg_end_t10/100
      ws.getCell("P"+countx10).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx10).value = 0/100
      ws.getCell("P"+countx10).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t10 > 0 && isNaN(this.sum_spliceloss_t10)==false && isFinite(this.sum_spliceloss_t10)==true){
      ws.getCell("R"+countx10).value = Number(this.sum_spliceloss_t10)
      ws.getCell("R"+countx10).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx10).value = 0.00
      ws.getCell("R"+countx10).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t10 > 0 && isNaN(this.sum_splicelossper_t10)==false && isFinite(this.sum_splicelossper_t10)==true){
      ws.getCell("S"+countx10).value = this.sum_splicelossper_t10/100
      ws.getCell("S"+countx10).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx10).value = 0/100
      ws.getCell("S"+countx10).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t10 > 0 && isNaN(this.sum_totalcutout_t10)==false && isFinite(this.sum_totalcutout_t10)==true){
      ws.getCell("U"+countx10).value = Number(this.sum_totalcutout_t10)
      ws.getCell("U"+countx10).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx10).value = 0.00
      ws.getCell("U"+countx10).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t10 > 0 && isNaN(this.sum_totalcutoutper_t10)==false && isFinite(this.sum_totalcutoutper_t10)==true){
      ws.getCell("V"+countx10).value = this.sum_totalcutoutper_t10/100
      ws.getCell("V"+countx10).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx10).value = 0.00/100
      ws.getCell("V"+countx10).numFmt = '0.00%';
      }

       if(this.sum_rement_t10 > 0 && isNaN(this.sum_rement_t10)==false && isFinite(this.sum_rement_t10)==true){
      ws.getCell("X"+countx10).value = Number(this.sum_rement_t10)
      ws.getCell("X"+countx10).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx10).value = 0.00
      ws.getCell("X"+countx10).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t10 > 0 && isNaN(this.sum_rement_per_t10)==false && isFinite(this.sum_rement_per_t10)==true){
      ws.getCell("Y"+countx10).value =  this.sum_rement_per_t10/100
      ws.getCell("Y"+countx10).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx10).value =  0.00/100
      ws.getCell("Y"+countx10).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t10 > 0 && isNaN(this.sum_totallenthloss_t10)==false && isFinite(this.sum_totallenthloss_t10)==true){
      ws.getCell("AA"+countx10).value = Number(this.sum_totallenthloss_t10)
      ws.getCell("AA"+countx10).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx10).value = 0
      ws.getCell("AA"+countx10).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t10 = this.sum_totallenthloss_t10 / this.sum_ttl_marker_t10 *100
      
      if(this.last_sum_totallenthloss_per_t10 > 0 && isNaN(this.last_sum_totallenthloss_per_t10)==false && isFinite(this.last_sum_totallenthloss_per_t10)==true){
      ws.getCell("AB"+countx10).value = this.last_sum_totallenthloss_per_t10/100
      ws.getCell("AB"+countx10).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx10).value = 0/100
      ws.getCell("AB"+countx10).numFmt = '0.00%'; 
      }

      
     

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx10).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //A11

      if(this.rowexport_type_10.length == 0){
       
         var count11 = count10

       }else{
      var count11 = countx10+1
       }
      
  
        var countx11 = count11 + 1 +Number(this.rowexport_type_11.length)
      
      this.sum_ttl_marker_t11 = 0
      this.sum_plug12_t11 = 0
      this.sum_per_width_loss_t11 = 0
      this.sum_end1_2_t11 = 0
      this.sum_endless_length_yd_t11= 0
      this.sum_avg_end_t11 = 0
      this.sum_spliceloss_t11 = 0
      this.sum_splicelossper_t11 = 0
      this.sum_totalcutout_t11 = 0
      this.sum_totalcutoutper_t11 = 0
      this.sum_rement_t11 = 0
      this.sum_rement_per_t11 = 0
      this.sum_totallenthloss_t11 =0 
      this.sum_totallenthlossper_t11 = 0
      this.summary_widthloss_t11 = 0
      this.summary_plug12_t11 = 0
      this.summary_end1_2_t11 = 0

      
      this.sum_all_plug1_2_end_total_t11 = 0

      if(this.rowexport_type_11.length < 1){
        count_row_grand++
      }else{
          for (var ax = 0; ax<this.rowexport_type_11.length; ax++){ //length +110
      ws.getCell("A"+count11).value = this.rowexport_type_11[ax].mu_date

      ws.getCell("B"+count11).value = this.rowexport_type_11[ax].gm_table

      ws.getCell("C"+count11).value = Number(this.rowexport_type_11[ax].gm_no_short)
      ws.getCell("C"+count11).numFmt = '0.00';

      ws.getCell("D"+count11).value = this.rowexport_type_11[ax].customer

      ws.getCell("E"+count11).value = this.rowexport_type_11[ax].gm_no

      ws.getCell("F"+count11).value = this.rowexport_type_11[ax].fabric_type

      ws.getCell("G"+count11).value = Number(this.rowexport_type_11[ax].obs_width)
      ws.getCell("G"+count11).numFmt = '0.00';

      ws.getCell("H"+count11).value = Number(this.rowexport_type_11[ax].width_inch)
      ws.getCell("H"+count11).numFmt = '0.00';


      ws.getCell("I"+count11).value = Number(this.rowexport_type_11[ax].length_ydb)
      ws.getCell("I"+count11).numFmt = '0.00';

      ws.getCell("J"+count11).value = Number(this.rowexport_type_11[ax].ttl_marker)
      ws.getCell("J"+count11).numFmt = '0.00';

      ws.getCell("K"+count11).value = Number(this.rowexport_type_11[ax].plug12)
      ws.getCell("K"+count11).numFmt = '0.00';

      ws.getCell("L"+count11).value = this.rowexport_type_11[ax].widthloss/100
      ws.getCell("L"+count11).numFmt = '0.00%';

      ws.getCell("M"+count11).value = this.rowexport_type_11[ax].widthloss_code

      ws.getCell("O"+count11).value = Number(this.rowexport_type_11[ax].endless_length_yd)
      ws.getCell("O"+count11).numFmt = '0.00';

      ws.getCell("P"+count11).value = this.rowexport_type_11[ax].avg_end/100
      ws.getCell("P"+count11).numFmt = '0.00%';

      ws.getCell("Q"+count11).value = this.rowexport_type_11[ax].avg_end_code
      

      ws.getCell("R"+count11).value = Number(this.rowexport_type_11[ax].spliceloss)
      ws.getCell("R"+count11).numFmt = '0.00';

      ws.getCell("S"+count11).value = this.rowexport_type_11[ax].splicelossper/100
      ws.getCell("S"+count11).numFmt = '0.00%';

      ws.getCell("T"+count11).value = this.rowexport_type_11[ax].per_splice_loss_code
   

      ws.getCell("U"+count11).value = Number(this.rowexport_type_11[ax].totalcutout)
      ws.getCell("U"+count11).numFmt = '0.00';

      ws.getCell("V"+count11).value = this.rowexport_type_11[ax].totalcutoutper/100
      ws.getCell("V"+count11).numFmt = '0.00%';

      ws.getCell("W"+count11).value = this.rowexport_type_11[ax].total_cut_out_per_code

      ws.getCell("X"+count11).value = Number(this.rowexport_type_11[ax].rement)
      ws.getCell("X"+count11).numFmt = '0.00';

      ws.getCell("Y"+count11).value = this.rowexport_type_11[ax].rement_per/100
      ws.getCell("Y"+count11).numFmt = '0.00%';

      ws.getCell("Z"+count11).value = this.rowexport_type_11[ax].remnants_loss_code

      ws.getCell("AA"+count11).value = Number(this.rowexport_type_11[ax].totallenthloss)
      ws.getCell("AA"+count11).numFmt = '0.00';

      ws.getCell("AB"+count11).value = this.rowexport_type_11[ax].totallenthlossper/100
      ws.getCell("AB"+count11).numFmt = '0.00%';
      
      this.sum_plug12_t11 = 0
      this.sum_ttl_marker_t11 = Number(this.sum_ttl_marker_t11) + Number(this.rowexport_type_11[ax].ttl_marker)
      this.sum_plug12_t11 = this.rowexport_type_11[ax].plug12 * this.rowexport_type_11[ax].ttl_marker
      this.sum_per_width_loss_t11 = this.rowexport_type_11[ax].widthloss * this.rowexport_type_11[ax].ttl_marker
      this.sum_end1_2_t11 = this.rowexport_type_11[ax].end1_2 * this.rowexport_type_11[ax].ttl_marker
      this.sum_endless_length_yd_t11 = Number(this.sum_endless_length_yd_t11) + Number(this.rowexport_type_11[ax].endless_length_yd)
      this.sum_spliceloss_t11 = Number(this.sum_spliceloss_t11) + Number(this.rowexport_type_11[ax].spliceloss)
      this.sum_totalcutout_t11 = Number(this.sum_totalcutout_t11) + Number(this.rowexport_type_11[ax].totalcutout)
      this.sum_rement_t11 = Number(this.sum_rement_t11) + Number(this.rowexport_type_11[ax].rement)
      this.summary_plug12_t11 = Number(this.summary_plug12_t11) + Number(this.sum_plug12_t11)
      this.summary_widthloss_t11 = Number(this.summary_widthloss_t11) + Number(this.sum_per_width_loss_t11)
      this.summary_end1_2_t11 = Number(this.summary_end1_2_t11) + Number(this.sum_end1_2_t11)
     
      this.sum_totallenthloss_t11_first = 0
      this.sum_totallenthloss_t11_first = Number(this.rowexport_type_11[ax].endless_length_yd) + Number(this.rowexport_type_11[ax].spliceloss)
      +Number(this.rowexport_type_11[ax].totalcutout) + Number(this.rowexport_type_11[ax].rement)
      this.sum_totallenthloss_t11 = Number(this.sum_totallenthloss_t11) + Number(this.sum_totallenthloss_t11_first)

      this.sum_totallenthloss_per_t11_first = 0
      this.sum_totallenthloss_per_t11_first = Number(this.rowexport_type_11[ax].avg_end) + Number(this.rowexport_type_11[ax].splicelossper)
      +Number(this.rowexport_type_11[ax].totalcutoutper) + Number(this.rowexport_type_11[ax].rement_per)
      this.sum_totallenthloss_per_t11 = Number(this.sum_totallenthloss_t11) + Number(this.sum_totallenthloss_per_t11_first)
      this.sum_number_1_end_t11 = 0 
      this.sum_number_1_end_t11 = this.rowexport_type_11[ax].endless_length_yd * 36
      this.sum_number_2_end_t11 = 0
      this.sum_number_2_end_t11 = this.rowexport_type_11[ax].ttl_marker/this.rowexport_type_11[ax].length_ydb
      this.sum_plug1_2_end_t11 = 0
      this.sum_plug1_2_end_t11 = this.sum_number_1_end_t11/this.sum_number_2_end_t11

      this.sum_plug1_2_end_total_t11 = 0
     
      this.sum_plug1_2_end_total_t11 = this.sum_plug1_2_end_t11 * this.rowexport_type_11[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t11 = Number(this.sum_all_plug1_2_end_total_t11) + Number(this.sum_plug1_2_end_total_t11)

      this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t11
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t11
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t11
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t11
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t11
      })

      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t11,
        value_table:11
      })


      ws.getCell("N"+count11).value = Number(this.sum_plug1_2_end_t11)
      ws.getCell("N"+count11).numFmt = '0.00';

      count11++
         }
         if(this.sum_ttl_marker_t11 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t11
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t11 = this.summary_plug12_t11 / this.sum_ttl_marker_t11
        if(this.sum_plug12_t11 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t11
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t11 =  this.summary_widthloss_t11 / this.sum_ttl_marker_t11
        if(this.sum_per_width_loss_t11 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t11
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t11 =  this.summary_end1_2_t11/ this.sum_ttl_marker_t11 //M
      if(this.sum_end1_2_t11 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t11
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t11 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t11
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t11 = this.sum_endless_length_yd_t11 / this.sum_ttl_marker_t11 *100 //P
      if(this.sum_avg_end_t11 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t11
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t11 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t11
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t11 = this.sum_spliceloss_t11 / this.sum_ttl_marker_t11 *100
      if(this.sum_splicelossper_t11 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t11
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t11 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t11
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t11 = this.sum_totalcutout_t11 / this.sum_ttl_marker_t11 *100
      if(this.sum_totalcutoutper_t11 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t11
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t11 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t11
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t11 = this.sum_rement_t11 / this.sum_ttl_marker_t11 *100
  if(this.sum_rement_per_t11 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t11
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t11
    if(this.sum_totallenthloss_t11 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t11
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert11 = this.sum_totallenthlosst11 / this.sum_ttl_marker_t11
    if(this.sum_totallenthlosspert11 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert11
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

ws.getCell("A"+countx11).value = "Total Table A11"  
      ws.mergeCells("A"+countx11+":"+"I"+countx11)
      
      if(this.sum_ttl_marker_t11 > 0 && isNaN(this.sum_ttl_marker_t11)==false && isFinite(this.sum_ttl_marker_t11)==true){
      ws.getCell("J"+countx11).value = Number(this.sum_ttl_marker_t11)
      ws.getCell("J"+countx11).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx11).value = 0
      ws.getCell("J"+countx11).numFmt = '0.00';
      }

      if(this.sum_plug12_t11 > 0 && isNaN(this.sum_plug12_t11)==false && isFinite(this.sum_plug12_t11)==true){
      ws.getCell("K"+countx11).value = Number(this.sum_plug12_t11)
      ws.getCell("K"+countx11).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx11).value = 0
      ws.getCell("K"+countx11).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t11 > 0 && isNaN(this.sum_per_width_loss_t11)==false && isFinite(this.sum_per_width_loss_t11)==true){
      ws.getCell("L"+countx11).value = this.sum_per_width_loss_t11/100
      ws.getCell("L"+countx11).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx11).value = 0.00/100
      ws.getCell("L"+countx11).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t11 = 0
      this.total_result_plug1_2_t11 = this.sum_all_plug1_2_end_total_t11 / this.sum_ttl_marker_t11

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t11
      })
      if(this.total_result_plug1_2_t11 > 0 && isNaN(this.total_result_plug1_2_t11)==false && isFinite(this.total_result_plug1_2_t11)==true){
      ws.getCell("N"+countx11).value = Number(this.total_result_plug1_2_t11)
      ws.getCell("N"+countx11).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx11).value = 0.00
      ws.getCell("N"+countx11).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_t11 > 0 && isNaN(this.sum_endless_length_yd_t11)==false && isFinite(this.sum_endless_length_yd_t11)==true){
      ws.getCell("O"+countx11).value = Number(this.sum_endless_length_yd_t11)
      ws.getCell("O"+countx11).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx11).value = 0.00
      ws.getCell("O"+countx11).numFmt = '0.00';
      }

      if(this.sum_avg_end_t11 > 0 && isNaN(this.sum_avg_end_t11)==false && isFinite(this.sum_avg_end_t11)==true){
      ws.getCell("P"+countx11).value = this.sum_avg_end_t11/100
      ws.getCell("P"+countx11).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx11).value = 0/100
      ws.getCell("P"+countx11).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t11 > 0 && isNaN(this.sum_spliceloss_t11)==false && isFinite(this.sum_spliceloss_t11)==true){
      ws.getCell("R"+countx11).value = Number(this.sum_spliceloss_t11)
      ws.getCell("R"+countx11).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx11).value = 0.00
      ws.getCell("R"+countx11).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t11 > 0 && isNaN(this.sum_splicelossper_t11)==false && isFinite(this.sum_splicelossper_t11)==true){
      ws.getCell("S"+countx11).value = this.sum_splicelossper_t11/100
      ws.getCell("S"+countx11).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx11).value = 0/100
      ws.getCell("S"+countx11).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t11 > 0 && isNaN(this.sum_totalcutout_t11)==false && isFinite(this.sum_totalcutout_t11)==true){
      ws.getCell("U"+countx11).value = Number(this.sum_totalcutout_t11)
      ws.getCell("U"+countx11).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx11).value = 0.00
      ws.getCell("U"+countx11).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t11 > 0 && isNaN(this.sum_totalcutoutper_t11)==false && isFinite(this.sum_totalcutoutper_t11)==true){
      ws.getCell("V"+countx11).value = this.sum_totalcutoutper_t11/100
      ws.getCell("V"+countx11).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx11).value = 0.00/100
      ws.getCell("V"+countx11).numFmt = '0.00%';
      }

       if(this.sum_rement_t11 > 0 && isNaN(this.sum_rement_t11)==false && isFinite(this.sum_rement_t11)==true){
      ws.getCell("X"+countx11).value = Number(this.sum_rement_t11)
      ws.getCell("X"+countx11).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx11).value = 0.00
      ws.getCell("X"+countx11).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t11 > 0 && isNaN(this.sum_rement_per_t11)==false && isFinite(this.sum_rement_per_t11)==true){
      ws.getCell("Y"+countx11).value =  this.sum_rement_per_t11/100
      ws.getCell("Y"+countx11).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx11).value =  0.00/100
      ws.getCell("Y"+countx11).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t11 > 0 && isNaN(this.sum_totallenthloss_t11)==false && isFinite(this.sum_totallenthloss_t11)==true){
      ws.getCell("AA"+countx11).value = Number(this.sum_totallenthloss_t11)
      ws.getCell("AA"+countx11).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx11).value = 0
      ws.getCell("AA"+countx11).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t11 = this.sum_totallenthloss_t11 / this.sum_ttl_marker_t11 *100
      if(this.last_sum_totallenthloss_per_t11 > 0 && isNaN(this.last_sum_totallenthloss_per_t11)==false && isFinite(this.last_sum_totallenthloss_per_t11)==true){
      ws.getCell("AB"+countx11).value = this.last_sum_totallenthloss_per_t11/100
      ws.getCell("AB"+countx11).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx11).value = 0/100
      ws.getCell("AB"+countx11).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx11).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //A12

      if(this.rowexport_type_11.length == 0){
       
         var count12 = count11

       }else{
      var count12 = countx11+1
       }
      
  
        var countx12 = count12 + 1 +Number(this.rowexport_type_12.length)
      
      this.sum_ttl_marker_t12 = 0
      this.sum_plug12_t12 = 0
      this.sum_per_width_loss_t12 = 0
      this.sum_end1_2_t12 = 0
      this.sum_endless_length_yd_t12= 0
      this.sum_avg_end_t12 = 0
      this.sum_spliceloss_t12 = 0
      this.sum_splicelossper_t12 = 0
      this.sum_totalcutout_t12 = 0
      this.sum_totalcutoutper_t12 = 0
      this.sum_rement_t12 = 0
      this.sum_rement_per_t12 = 0
      this.sum_totallenthloss_t12 =0 
      this.sum_totallenthlossper_t12 = 0
      this.summary_widthloss_t12 = 0
      this.summary_plug12_t12 = 0
      this.summary_end1_2_t12 = 0

      
     
      this.sum_all_plug1_2_end_total_t12 = 0

      if(this.rowexport_type_12.length < 1){
        count_row_grand++
      }else{
          for (var ax = 0; ax<this.rowexport_type_12.length; ax++){ //length +120
      ws.getCell("A"+count12).value = this.rowexport_type_12[ax].mu_date

      ws.getCell("B"+count12).value = this.rowexport_type_12[ax].gm_table

      ws.getCell("C"+count12).value = Number(this.rowexport_type_12[ax].gm_no_short)
      ws.getCell("C"+count12).numFmt = '0.00';

      ws.getCell("D"+count12).value = this.rowexport_type_12[ax].customer

      ws.getCell("E"+count12).value = this.rowexport_type_12[ax].gm_no

      ws.getCell("F"+count12).value = this.rowexport_type_12[ax].fabric_type

      ws.getCell("G"+count12).value = Number(this.rowexport_type_12[ax].obs_width)
      ws.getCell("G"+count12).numFmt = '0.00';

      ws.getCell("H"+count12).value = Number(this.rowexport_type_12[ax].width_inch)
      ws.getCell("H"+count12).numFmt = '0.00';


      ws.getCell("I"+count12).value = Number(this.rowexport_type_12[ax].length_ydb)
      ws.getCell("I"+count12).numFmt = '0.00';

      ws.getCell("J"+count12).value = Number(this.rowexport_type_12[ax].ttl_marker)
      ws.getCell("J"+count12).numFmt = '0.00';

      ws.getCell("K"+count12).value = Number(this.rowexport_type_12[ax].plug12)
      ws.getCell("K"+count12).numFmt = '0.00';

      ws.getCell("L"+count12).value = this.rowexport_type_12[ax].widthloss/100
      ws.getCell("L"+count12).numFmt = '0.00%';

      ws.getCell("M"+count12).value = this.rowexport_type_12[ax].widthloss_code
    
      ws.getCell("O"+count12).value = Number(this.rowexport_type_12[ax].endless_length_yd)
      ws.getCell("O"+count12).numFmt = '0.00';

      ws.getCell("P"+count12).value = this.rowexport_type_12[ax].avg_end/100
      ws.getCell("P"+count12).numFmt = '0.00%';

      ws.getCell("Q"+count12).value = this.rowexport_type_12[ax].avg_end_code
      

      ws.getCell("R"+count12).value = Number(this.rowexport_type_12[ax].spliceloss)
      ws.getCell("R"+count12).numFmt = '0.00';

      ws.getCell("S"+count12).value = this.rowexport_type_12[ax].splicelossper/100
      ws.getCell("S"+count12).numFmt = '0.00%';

      ws.getCell("T"+count12).value = this.rowexport_type_12[ax].per_splice_loss_code
   

      ws.getCell("U"+count12).value = Number(this.rowexport_type_12[ax].totalcutout)
      ws.getCell("U"+count12).numFmt = '0.00';

      ws.getCell("V"+count12).value = this.rowexport_type_12[ax].totalcutoutper/100
      ws.getCell("V"+count12).numFmt = '0.00%';

      ws.getCell("W"+count12).value = this.rowexport_type_12[ax].total_cut_out_per_code

      ws.getCell("X"+count12).value = Number(this.rowexport_type_12[ax].rement)
      ws.getCell("X"+count12).numFmt = '0.00';

      ws.getCell("Y"+count12).value = this.rowexport_type_12[ax].rement_per/100
      ws.getCell("Y"+count12).numFmt = '0.00%';

      ws.getCell("Z"+count12).value = this.rowexport_type_12[ax].remnants_loss_code

      ws.getCell("AA"+count12).value = Number(this.rowexport_type_12[ax].totallenthloss)
      ws.getCell("AA"+count12).numFmt = '0.00';

      ws.getCell("AB"+count12).value = this.rowexport_type_12[ax].totallenthlossper/100
      ws.getCell("AB"+count12).numFmt = '0.00%';
      
      this.sum_plug12_t12 = 0
      this.sum_ttl_marker_t12 = Number(this.sum_ttl_marker_t12) + Number(this.rowexport_type_12[ax].ttl_marker)
      this.sum_plug12_t12 = this.rowexport_type_12[ax].plug12 * this.rowexport_type_12[ax].ttl_marker
      this.sum_per_width_loss_t12 = this.rowexport_type_12[ax].widthloss * this.rowexport_type_12[ax].ttl_marker
      this.sum_end1_2_t12 = this.rowexport_type_12[ax].end1_2 * this.rowexport_type_12[ax].ttl_marker
      this.sum_endless_length_yd_t12 = Number(this.sum_endless_length_yd_t12) + Number(this.rowexport_type_12[ax].endless_length_yd)
      this.sum_spliceloss_t12 = Number(this.sum_spliceloss_t12) + Number(this.rowexport_type_12[ax].spliceloss)
      this.sum_totalcutout_t12 = Number(this.sum_totalcutout_t12) + Number(this.rowexport_type_12[ax].totalcutout)
      this.sum_rement_t12 = Number(this.sum_rement_t12) + Number(this.rowexport_type_12[ax].rement)
      this.summary_plug12_t12 = Number(this.summary_plug12_t12) + Number(this.sum_plug12_t12)
      this.summary_widthloss_t12 = Number(this.summary_widthloss_t12) + Number(this.sum_per_width_loss_t12)
      this.summary_end1_2_t12 = Number(this.summary_end1_2_t12) + Number(this.sum_end1_2_t12)
      this.sum_totallenthloss_t12_first = 0
      this.sum_totallenthloss_t12_first = Number(this.rowexport_type_12[ax].endless_length_yd) + Number(this.rowexport_type_12[ax].spliceloss)
      +Number(this.rowexport_type_12[ax].totalcutout) + Number(this.rowexport_type_12[ax].rement)
      this.sum_totallenthloss_t12 = Number(this.sum_totallenthloss_t12) + Number(this.sum_totallenthloss_t12_first)

      this.sum_totallenthloss_per_t12_first = 0
      this.sum_totallenthloss_per_t12_first = Number(this.rowexport_type_12[ax].avg_end) + Number(this.rowexport_type_12[ax].splicelossper)
      +Number(this.rowexport_type_12[ax].totalcutoutper) + Number(this.rowexport_type_12[ax].rement_per)
      this.sum_totallenthloss_per_t12 = Number(this.sum_totallenthloss_t12) + Number(this.sum_totallenthloss_per_t12_first)
      this.sum_number_1_end_t12 = 0 
      this.sum_number_1_end_t12 = this.rowexport_type_12[ax].endless_length_yd * 36
      this.sum_number_2_end_t12 = 0
      this.sum_number_2_end_t12 = this.rowexport_type_12[ax].ttl_marker/this.rowexport_type_12[ax].length_ydb
      this.sum_plug1_2_end_t12 = 0
      this.sum_plug1_2_end_t12 = this.sum_number_1_end_t12/this.sum_number_2_end_t12

      this.sum_plug1_2_end_total_t12 = 0
     
      this.sum_plug1_2_end_total_t12 = this.sum_plug1_2_end_t12 * this.rowexport_type_12[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t12 = Number(this.sum_all_plug1_2_end_total_t12) + Number(this.sum_plug1_2_end_total_t12)

       this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t12
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t12
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t12
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t12
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t12
      })
       
      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t12,
        value_table:12
      }) 

      ws.getCell("N"+count12).value = Number(this.sum_plug1_2_end_t12)
      ws.getCell("N"+count12).numFmt = '0.00';


      count12++
         }
         if(this.sum_ttl_marker_t12 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t12
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t12 = this.summary_plug12_t12 / this.sum_ttl_marker_t12
        if(this.sum_plug12_t12 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t12
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t12 =  this.summary_widthloss_t12 / this.sum_ttl_marker_t12
        if(this.sum_per_width_loss_t12 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t12
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t12 =  this.summary_end1_2_t12/ this.sum_ttl_marker_t12 //M
      if(this.sum_end1_2_t12 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t12
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t12 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t12
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t12 = this.sum_endless_length_yd_t12 / this.sum_ttl_marker_t12 *100 //P
      if(this.sum_avg_end_t12 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t12
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t12 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t12
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t12 = this.sum_spliceloss_t12 / this.sum_ttl_marker_t12 *100
      if(this.sum_splicelossper_t12 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t12
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t12 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t12
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t12 = this.sum_totalcutout_t12 / this.sum_ttl_marker_t12 *100
      if(this.sum_totalcutoutper_t12 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t12
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t12 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t12
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t12 = this.sum_rement_t12 / this.sum_ttl_marker_t12 *100
  if(this.sum_rement_per_t12 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t12
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t12
    if(this.sum_totallenthloss_t12 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t12
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert12 = this.sum_totallenthlosst12 / this.sum_ttl_marker_t12
    if(this.sum_totallenthlosspert12 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert12
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

ws.getCell("A"+countx12).value = "Total Table A12"  
      ws.mergeCells("A"+countx12+":"+"I"+countx12)
      
      if(this.sum_ttl_marker_t12 > 0 && isNaN(this.sum_ttl_marker_t12)==false && isFinite(this.sum_ttl_marker_t12)==true){
      ws.getCell("J"+countx12).value = Number(this.sum_ttl_marker_t12)
      ws.getCell("J"+countx12).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx12).value = 0
      ws.getCell("J"+countx12).numFmt = '0.00';
      }

      if(this.sum_plug12_t12 > 0 && isNaN(this.sum_plug12_t12)==false && isFinite(this.sum_plug12_t12)==true){
      ws.getCell("K"+countx12).value = Number(this.sum_plug12_t12)
      ws.getCell("K"+countx12).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx12).value = 0
      ws.getCell("K"+countx12).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t12 > 0 && isNaN(this.sum_per_width_loss_t12)==false && isFinite(this.sum_per_width_loss_t12)==true){
      ws.getCell("L"+countx12).value = this.sum_per_width_loss_t12/100
      ws.getCell("L"+countx12).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx12).value = 0.00/100
      ws.getCell("L"+countx12).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t12 = 0
      this.total_result_plug1_2_t12 = this.sum_all_plug1_2_end_total_t12 / this.sum_ttl_marker_t12

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t12
      })
      if(this.total_result_plug1_2_t12 > 0 && isNaN(this.total_result_plug1_2_t12)==false && isFinite(this.total_result_plug1_2_t12)==true){
      ws.getCell("N"+countx12).value = Number(this.total_result_plug1_2_t12)
      ws.getCell("N"+countx12).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx12).value = 0.00
      ws.getCell("N"+countx12).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_t12 > 0 && isNaN(this.sum_endless_length_yd_t12)==false && isFinite(this.sum_endless_length_yd_t12)==true){
      ws.getCell("O"+countx12).value = Number(this.sum_endless_length_yd_t12)
      ws.getCell("O"+countx12).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx12).value = 0.00
      ws.getCell("O"+countx12).numFmt = '0.00';
      }

      if(this.sum_avg_end_t12 > 0 && isNaN(this.sum_avg_end_t12)==false && isFinite(this.sum_avg_end_t12)==true){
      ws.getCell("P"+countx12).value = this.sum_avg_end_t12/100
      ws.getCell("P"+countx12).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx12).value = 0/100
      ws.getCell("P"+countx12).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t12 > 0 && isNaN(this.sum_spliceloss_t12)==false && isFinite(this.sum_spliceloss_t12)==true){
      ws.getCell("R"+countx12).value = Number(this.sum_spliceloss_t12)
      ws.getCell("R"+countx12).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx12).value = 0.00
      ws.getCell("R"+countx12).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t12 > 0 && isNaN(this.sum_splicelossper_t12)==false && isFinite(this.sum_splicelossper_t12)==true){
      ws.getCell("S"+countx12).value = this.sum_splicelossper_t12/100
      ws.getCell("S"+countx12).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx12).value = 0/100
      ws.getCell("S"+countx12).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t12 > 0 && isNaN(this.sum_totalcutout_t12)==false && isFinite(this.sum_totalcutout_t12)==true){
      ws.getCell("U"+countx12).value = Number(this.sum_totalcutout_t12)
      ws.getCell("U"+countx12).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx12).value = 0.00
      ws.getCell("U"+countx12).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t12 > 0 && isNaN(this.sum_totalcutoutper_t12)==false && isFinite(this.sum_totalcutoutper_t12)==true){
      ws.getCell("V"+countx12).value = this.sum_totalcutoutper_t12/100
      ws.getCell("V"+countx12).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx12).value = 0.00/100
      ws.getCell("V"+countx12).numFmt = '0.00%';
      }

       if(this.sum_rement_t12 > 0 && isNaN(this.sum_rement_t12)==false && isFinite(this.sum_rement_t12)==true){
      ws.getCell("X"+countx12).value = Number(this.sum_rement_t12)
      ws.getCell("X"+countx12).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx12).value = 0.00
      ws.getCell("X"+countx12).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t12 > 0 && isNaN(this.sum_rement_per_t12)==false && isFinite(this.sum_rement_per_t12)==true){
      ws.getCell("Y"+countx12).value =  this.sum_rement_per_t12/100
      ws.getCell("Y"+countx12).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx12).value =  0.00/100
      ws.getCell("Y"+countx12).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t12 > 0 && isNaN(this.sum_totallenthloss_t12)==false && isFinite(this.sum_totallenthloss_t12)==true){
      ws.getCell("AA"+countx12).value = Number(this.sum_totallenthloss_t12)
      ws.getCell("AA"+countx12).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx12).value = 0
      ws.getCell("AA"+countx12).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t12 = this.sum_totallenthloss_t12 / this.sum_ttl_marker_t12 *100
      if(this.last_sum_totallenthloss_per_t12 > 0 && isNaN(this.last_sum_totallenthloss_per_t12)==false && isFinite(this.last_sum_totallenthloss_per_t12)==true){
      ws.getCell("AB"+countx12).value = this.last_sum_totallenthloss_per_t12/100
      ws.getCell("AB"+countx12).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx12).value = 0/100
      ws.getCell("AB"+countx12).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx12).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }


    if(this.rowexport_type_12.length == 0){
       
         var count13 = count12

       }else{
      var count13 = countx12+1
       }
      
      
  
        var countx13 = count13 + 1 +Number(this.rowexport_type_13.length)
      
      this.sum_ttl_marker_t13 = 0
      this.sum_plug12_t13 = 0
      this.sum_per_width_loss_t13 = 0
      this.sum_end1_2_t13 = 0
      this.sum_endless_length_yd_t13= 0
      this.sum_avg_end_t13 = 0
      this.sum_spliceloss_t13 = 0
      this.sum_splicelossper_t13 = 0
      this.sum_totalcutout_t13 = 0
      this.sum_totalcutoutper_t13 = 0
      this.sum_rement_t13 = 0
      this.sum_rement_per_t13 = 0
      this.sum_totallenthloss_t13 =0 
      this.sum_totallenthlossper_t13 = 0
      this.summary_widthloss_t13 = 0
      this.summary_plug12_t13 = 0
      this.summary_end1_2_t13 = 0

      
     
      this.sum_all_plug1_2_end_total_t13 = 0

      if(this.rowexport_type_13.length < 1){
        count_row_grand++
      }else{
          for (var ax = 0; ax<this.rowexport_type_13.length; ax++){ //length +120
      ws.getCell("A"+count13).value = this.rowexport_type_13[ax].mu_date

      ws.getCell("B"+count13).value = this.rowexport_type_13[ax].gm_table

      ws.getCell("C"+count13).value = Number(this.rowexport_type_13[ax].gm_no_short)
      ws.getCell("C"+count13).numFmt = '0.00';

      ws.getCell("D"+count13).value = this.rowexport_type_13[ax].customer

      ws.getCell("E"+count13).value = this.rowexport_type_13[ax].gm_no

      ws.getCell("F"+count13).value = this.rowexport_type_13[ax].fabric_type

      ws.getCell("G"+count13).value = Number(this.rowexport_type_13[ax].obs_width)
      ws.getCell("G"+count13).numFmt = '0.00';

      ws.getCell("H"+count13).value = Number(this.rowexport_type_13[ax].width_inch)
      ws.getCell("H"+count13).numFmt = '0.00';


      ws.getCell("I"+count13).value = Number(this.rowexport_type_13[ax].length_ydb)
      ws.getCell("I"+count13).numFmt = '0.00';

      ws.getCell("J"+count13).value = Number(this.rowexport_type_13[ax].ttl_marker)
      ws.getCell("J"+count13).numFmt = '0.00';

      ws.getCell("K"+count13).value = Number(this.rowexport_type_13[ax].plug12)
      ws.getCell("K"+count13).numFmt = '0.00';

      ws.getCell("L"+count13).value = this.rowexport_type_13[ax].widthloss/100
      ws.getCell("L"+count13).numFmt = '0.00%';

      ws.getCell("M"+count13).value = this.rowexport_type_13[ax].widthloss_code
    
      ws.getCell("O"+count13).value = Number(this.rowexport_type_13[ax].endless_length_yd)
      ws.getCell("O"+count13).numFmt = '0.00';

      ws.getCell("P"+count13).value = this.rowexport_type_13[ax].avg_end/100
      ws.getCell("P"+count13).numFmt = '0.00%';

      ws.getCell("Q"+count13).value = this.rowexport_type_13[ax].avg_end_code
      

      ws.getCell("R"+count13).value = Number(this.rowexport_type_13[ax].spliceloss)
      ws.getCell("R"+count13).numFmt = '0.00';

      ws.getCell("S"+count13).value = this.rowexport_type_13[ax].splicelossper/100
      ws.getCell("S"+count13).numFmt = '0.00%';

      ws.getCell("T"+count13).value = this.rowexport_type_13[ax].per_splice_loss_code
   

      ws.getCell("U"+count13).value = Number(this.rowexport_type_13[ax].totalcutout)
      ws.getCell("U"+count13).numFmt = '0.00';

      ws.getCell("V"+count13).value = this.rowexport_type_13[ax].totalcutoutper/100
      ws.getCell("V"+count13).numFmt = '0.00%';

      ws.getCell("W"+count13).value = this.rowexport_type_13[ax].total_cut_out_per_code

      ws.getCell("X"+count13).value = Number(this.rowexport_type_13[ax].rement)
      ws.getCell("X"+count13).numFmt = '0.00';

      ws.getCell("Y"+count13).value = this.rowexport_type_13[ax].rement_per/100
      ws.getCell("Y"+count13).numFmt = '0.00%';

      ws.getCell("Z"+count13).value = this.rowexport_type_13[ax].remnants_loss_code

      ws.getCell("AA"+count13).value = Number(this.rowexport_type_13[ax].totallenthloss)
      ws.getCell("AA"+count13).numFmt = '0.00';

      ws.getCell("AB"+count13).value = this.rowexport_type_13[ax].totallenthlossper/100
      ws.getCell("AB"+count13).numFmt = '0.00%';
      
      this.sum_plug12_t13 = 0
      this.sum_ttl_marker_t13 = Number(this.sum_ttl_marker_t13) + Number(this.rowexport_type_13[ax].ttl_marker)
      this.sum_plug12_t13 = this.rowexport_type_13[ax].plug12 * this.rowexport_type_13[ax].ttl_marker
      this.sum_per_width_loss_t13 = this.rowexport_type_13[ax].widthloss * this.rowexport_type_13[ax].ttl_marker
      this.sum_end1_2_t13 = this.rowexport_type_13[ax].end1_2 * this.rowexport_type_13[ax].ttl_marker
      this.sum_endless_length_yd_t13 = Number(this.sum_endless_length_yd_t13) + Number(this.rowexport_type_13[ax].endless_length_yd)
      this.sum_spliceloss_t13 = Number(this.sum_spliceloss_t13) + Number(this.rowexport_type_13[ax].spliceloss)
      this.sum_totalcutout_t13 = Number(this.sum_totalcutout_t13) + Number(this.rowexport_type_13[ax].totalcutout)
      this.sum_rement_t13 = Number(this.sum_rement_t13) + Number(this.rowexport_type_13[ax].rement)
      this.summary_plug12_t13 = Number(this.summary_plug12_t13) + Number(this.sum_plug12_t13)
      this.summary_widthloss_t13 = Number(this.summary_widthloss_t13) + Number(this.sum_per_width_loss_t13)
      this.summary_end1_2_t13 = Number(this.summary_end1_2_t13) + Number(this.sum_end1_2_t13)
      this.sum_totallenthloss_t13_first = 0
      this.sum_totallenthloss_t13_first = Number(this.rowexport_type_13[ax].endless_length_yd) + Number(this.rowexport_type_13[ax].spliceloss)
      +Number(this.rowexport_type_13[ax].totalcutout) + Number(this.rowexport_type_13[ax].rement)
      this.sum_totallenthloss_t13 = Number(this.sum_totallenthloss_t13) + Number(this.sum_totallenthloss_t13_first)

      this.sum_totallenthloss_per_t13_first = 0
      this.sum_totallenthloss_per_t13_first = Number(this.rowexport_type_13[ax].avg_end) + Number(this.rowexport_type_13[ax].splicelossper)
      +Number(this.rowexport_type_13[ax].totalcutoutper) + Number(this.rowexport_type_13[ax].rement_per)
      this.sum_totallenthloss_per_t13 = Number(this.sum_totallenthloss_t13) + Number(this.sum_totallenthloss_per_t13_first)
      this.sum_number_1_end_t13 = 0 
      this.sum_number_1_end_t13 = this.rowexport_type_13[ax].endless_length_yd * 36
      this.sum_number_2_end_t13 = 0
      this.sum_number_2_end_t13 = this.rowexport_type_13[ax].ttl_marker/this.rowexport_type_13[ax].length_ydb
      this.sum_plug1_2_end_t13 = 0
      this.sum_plug1_2_end_t13 = this.sum_number_1_end_t13/this.sum_number_2_end_t13

      this.sum_plug1_2_end_total_t13 = 0
     
      this.sum_plug1_2_end_total_t13 = this.sum_plug1_2_end_t13 * this.rowexport_type_13[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t13 = Number(this.sum_all_plug1_2_end_total_t13) + Number(this.sum_plug1_2_end_total_t13)

       this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t13
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t13
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t13
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t13
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t13
      })
       
      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t13,
        value_table:13
      }) 

      ws.getCell("N"+count13).value = Number(this.sum_plug1_2_end_t13)
      ws.getCell("N"+count13).numFmt = '0.00';


      count13++
         }
         if(this.sum_ttl_marker_t13 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t13
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t13 = this.summary_plug12_t13 / this.sum_ttl_marker_t13
        if(this.sum_plug12_t13 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t13
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t13 =  this.summary_widthloss_t13 / this.sum_ttl_marker_t13
        if(this.sum_per_width_loss_t13 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t13
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t13 =  this.summary_end1_2_t13/ this.sum_ttl_marker_t13 //M
      if(this.sum_end1_2_t13 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t13
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t13 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t13
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t13 = this.sum_endless_length_yd_t13 / this.sum_ttl_marker_t13 *100 //P
      if(this.sum_avg_end_t13 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t13
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t13 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t13
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t13 = this.sum_spliceloss_t13 / this.sum_ttl_marker_t13 *100
      if(this.sum_splicelossper_t13 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t13
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t13 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t13
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t13 = this.sum_totalcutout_t13 / this.sum_ttl_marker_t13 *100
      if(this.sum_totalcutoutper_t13 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t13
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    console.log(this.sum_rement_t13)
    if(this.sum_rement_t13 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t13
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }
  console.log(this.sum_rement_t13)
  console.log(this.sum_ttl_marker_t13)
  console.log(this.sum_rement_per_t13)
  this.sum_rement_per_t13 = this.sum_rement_t13 / this.sum_ttl_marker_t13 *100
  if(this.sum_rement_per_t13 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t13
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t13
    if(this.sum_totallenthloss_t13 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t13
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert13 = this.sum_totallenthlosst13 / this.sum_ttl_marker_t13
    if(this.sum_totallenthlosspert13 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert13
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

ws.getCell("A"+countx13).value = "Total Table A13"  
      ws.mergeCells("A"+countx13+":"+"I"+countx13)
      
      if(this.sum_ttl_marker_t13 > 0 && isNaN(this.sum_ttl_marker_t13)==false && isFinite(this.sum_ttl_marker_t13)==true){
      ws.getCell("J"+countx13).value = Number(this.sum_ttl_marker_t13)
      ws.getCell("J"+countx13).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx13).value = 0
      ws.getCell("J"+countx13).numFmt = '0.00';
      }

      if(this.sum_plug12_t13 > 0 && isNaN(this.sum_plug12_t13)==false && isFinite(this.sum_plug12_t13)==true){
      ws.getCell("K"+countx13).value = Number(this.sum_plug12_t13)
      ws.getCell("K"+countx13).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx13).value = 0
      ws.getCell("K"+countx13).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t13 > 0 && isNaN(this.sum_per_width_loss_t13)==false && isFinite(this.sum_per_width_loss_t13)==true){
      ws.getCell("L"+countx13).value = this.sum_per_width_loss_t13/100
      ws.getCell("L"+countx13).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx13).value = 0.00/100
      ws.getCell("L"+countx13).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t13 = 0
      this.total_result_plug1_2_t13 = this.sum_all_plug1_2_end_total_t13 / this.sum_ttl_marker_t13

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t13
      })
      if(this.total_result_plug1_2_t13 > 0 && isNaN(this.total_result_plug1_2_t13)==false && isFinite(this.total_result_plug1_2_t13)==true){
      ws.getCell("N"+countx13).value = Number(this.total_result_plug1_2_t13)
      ws.getCell("N"+countx13).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx13).value = 0.00
      ws.getCell("N"+countx13).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_t13 > 0 && isNaN(this.sum_endless_length_yd_t13)==false && isFinite(this.sum_endless_length_yd_t13)==true){
      ws.getCell("O"+countx13).value = Number(this.sum_endless_length_yd_t13)
      ws.getCell("O"+countx13).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx13).value = 0.00
      ws.getCell("O"+countx13).numFmt = '0.00';
      }

      if(this.sum_avg_end_t13 > 0 && isNaN(this.sum_avg_end_t13)==false && isFinite(this.sum_avg_end_t13)==true){
      ws.getCell("P"+countx13).value = this.sum_avg_end_t13/100
      ws.getCell("P"+countx13).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx13).value = 0/100
      ws.getCell("P"+countx13).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t13 > 0 && isNaN(this.sum_spliceloss_t13)==false && isFinite(this.sum_spliceloss_t13)==true){
      ws.getCell("R"+countx13).value = Number(this.sum_spliceloss_t13)
      ws.getCell("R"+countx13).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx13).value = 0.00
      ws.getCell("R"+countx13).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t13 > 0 && isNaN(this.sum_splicelossper_t13)==false && isFinite(this.sum_splicelossper_t13)==true){
      ws.getCell("S"+countx13).value = this.sum_splicelossper_t13/100
      ws.getCell("S"+countx13).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx13).value = 0/100
      ws.getCell("S"+countx13).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t13 > 0 && isNaN(this.sum_totalcutout_t13)==false && isFinite(this.sum_totalcutout_t13)==true){
      ws.getCell("U"+countx13).value = Number(this.sum_totalcutout_t13)
      ws.getCell("U"+countx13).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx13).value = 0.00
      ws.getCell("U"+countx13).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t13 > 0 && isNaN(this.sum_totalcutoutper_t13)==false && isFinite(this.sum_totalcutoutper_t13)==true){
      ws.getCell("V"+countx13).value = this.sum_totalcutoutper_t13/100
      ws.getCell("V"+countx13).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx13).value = 0.00/100
      ws.getCell("V"+countx13).numFmt = '0.00%';
      }

       if(this.sum_rement_t13 > 0 && isNaN(this.sum_rement_t13)==false && isFinite(this.sum_rement_t13)==true){
      ws.getCell("X"+countx13).value = Number(this.sum_rement_t13)
      ws.getCell("X"+countx13).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx13).value = 0.00
      ws.getCell("X"+countx13).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t13 > 0 && isNaN(this.sum_rement_per_t13)==false && isFinite(this.sum_rement_per_t13)==true){
      ws.getCell("Y"+countx13).value =  this.sum_rement_per_t13/100
      ws.getCell("Y"+countx13).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx13).value =  0.00/100
      ws.getCell("Y"+countx13).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t13 > 0 && isNaN(this.sum_totallenthloss_t13)==false && isFinite(this.sum_totallenthloss_t13)==true){
      ws.getCell("AA"+countx13).value = Number(this.sum_totallenthloss_t13)
      ws.getCell("AA"+countx13).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx13).value = 0
      ws.getCell("AA"+countx13).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t13 = this.sum_totallenthloss_t13 / this.sum_ttl_marker_t13 *100
      if(this.last_sum_totallenthloss_per_t13 > 0 && isNaN(this.last_sum_totallenthloss_per_t13)==false && isFinite(this.last_sum_totallenthloss_per_t13)==true){
      ws.getCell("AB"+countx13).value = this.last_sum_totallenthloss_per_t13/100
      ws.getCell("AB"+countx13).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx13).value = 0/100
      ws.getCell("AB"+countx13).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx13).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }


      //A14


      if(this.rowexport_type_13.length == 0){
       
         var count14 = count13

       }else{
      var count14 = countx13+1
       }
      
  
        var countx14 = count14 + 1 +Number(this.rowexport_type_14.length)
      
      this.sum_ttl_marker_t14 = 0
      this.sum_plug12_t14 = 0
      this.sum_per_width_loss_t14 = 0
      this.sum_end1_2_t14 = 0
      this.sum_endless_length_yd_t14= 0
      this.sum_avg_end_t14 = 0
      this.sum_spliceloss_t14 = 0
      this.sum_splicelossper_t14 = 0
      this.sum_totalcutout_t14 = 0
      this.sum_totalcutoutper_t14 = 0
      this.sum_rement_t14 = 0
      this.sum_rement_per_t14 = 0
      this.sum_totallenthloss_t14 =0 
      this.sum_totallenthlossper_t14 = 0
      this.summary_widthloss_t14 = 0
      this.summary_plug12_t14 = 0
      this.summary_end1_2_t14 = 0

      
     
      this.sum_all_plug1_2_end_total_t14 = 0

      if(this.rowexport_type_14.length < 1){
        count_row_grand++
      }else{
          for (var ax = 0; ax<this.rowexport_type_14.length; ax++){ //length +120
      ws.getCell("A"+count14).value = this.rowexport_type_14[ax].mu_date

      ws.getCell("B"+count14).value = this.rowexport_type_14[ax].gm_table

      ws.getCell("C"+count14).value = Number(this.rowexport_type_14[ax].gm_no_short)
      ws.getCell("C"+count14).numFmt = '0.00';

      ws.getCell("D"+count14).value = this.rowexport_type_14[ax].customer

      ws.getCell("E"+count14).value = this.rowexport_type_14[ax].gm_no

      ws.getCell("F"+count14).value = this.rowexport_type_14[ax].fabric_type

      ws.getCell("G"+count14).value = Number(this.rowexport_type_14[ax].obs_width)
      ws.getCell("G"+count14).numFmt = '0.00';

      ws.getCell("H"+count14).value = Number(this.rowexport_type_14[ax].width_inch)
      ws.getCell("H"+count14).numFmt = '0.00';


      ws.getCell("I"+count14).value = Number(this.rowexport_type_14[ax].length_ydb)
      ws.getCell("I"+count14).numFmt = '0.00';

      ws.getCell("J"+count14).value = Number(this.rowexport_type_14[ax].ttl_marker)
      ws.getCell("J"+count14).numFmt = '0.00';

      ws.getCell("K"+count14).value = Number(this.rowexport_type_14[ax].plug12)
      ws.getCell("K"+count14).numFmt = '0.00';

      ws.getCell("L"+count14).value = this.rowexport_type_14[ax].widthloss/100
      ws.getCell("L"+count14).numFmt = '0.00%';

      ws.getCell("M"+count14).value = this.rowexport_type_14[ax].widthloss_code
    
      ws.getCell("O"+count14).value = Number(this.rowexport_type_14[ax].endless_length_yd)
      ws.getCell("O"+count14).numFmt = '0.00';

      ws.getCell("P"+count14).value = this.rowexport_type_14[ax].avg_end/100
      ws.getCell("P"+count14).numFmt = '0.00%';

      ws.getCell("Q"+count14).value = this.rowexport_type_14[ax].avg_end_code
      

      ws.getCell("R"+count14).value = Number(this.rowexport_type_14[ax].spliceloss)
      ws.getCell("R"+count14).numFmt = '0.00';

      ws.getCell("S"+count14).value = this.rowexport_type_14[ax].splicelossper/100
      ws.getCell("S"+count14).numFmt = '0.00%';

      ws.getCell("T"+count14).value = this.rowexport_type_14[ax].per_splice_loss_code
   

      ws.getCell("U"+count14).value = Number(this.rowexport_type_14[ax].totalcutout)
      ws.getCell("U"+count14).numFmt = '0.00';

      ws.getCell("V"+count14).value = this.rowexport_type_14[ax].totalcutoutper/100
      ws.getCell("V"+count14).numFmt = '0.00%';

      ws.getCell("W"+count14).value = this.rowexport_type_14[ax].total_cut_out_per_code

      ws.getCell("X"+count14).value = Number(this.rowexport_type_14[ax].rement)
      ws.getCell("X"+count14).numFmt = '0.00';

      ws.getCell("Y"+count14).value = this.rowexport_type_14[ax].rement_per/100
      ws.getCell("Y"+count14).numFmt = '0.00%';

      ws.getCell("Z"+count14).value = this.rowexport_type_14[ax].remnants_loss_code

      ws.getCell("AA"+count14).value = Number(this.rowexport_type_14[ax].totallenthloss)
      ws.getCell("AA"+count14).numFmt = '0.00';

      ws.getCell("AB"+count14).value = this.rowexport_type_14[ax].totallenthlossper/100
      ws.getCell("AB"+count14).numFmt = '0.00%';
      
      this.sum_plug12_t14 = 0
      this.sum_ttl_marker_t14 = Number(this.sum_ttl_marker_t14) + Number(this.rowexport_type_14[ax].ttl_marker)
      this.sum_plug12_t14 = this.rowexport_type_14[ax].plug12 * this.rowexport_type_14[ax].ttl_marker
      this.sum_per_width_loss_t14 = this.rowexport_type_14[ax].widthloss * this.rowexport_type_14[ax].ttl_marker
      this.sum_end1_2_t14 = this.rowexport_type_14[ax].end1_2 * this.rowexport_type_14[ax].ttl_marker
      this.sum_endless_length_yd_t14 = Number(this.sum_endless_length_yd_t14) + Number(this.rowexport_type_14[ax].endless_length_yd)
      this.sum_spliceloss_t14 = Number(this.sum_spliceloss_t14) + Number(this.rowexport_type_14[ax].spliceloss)
      this.sum_totalcutout_t14 = Number(this.sum_totalcutout_t14) + Number(this.rowexport_type_14[ax].totalcutout)
      this.sum_rement_t14 = Number(this.sum_rement_t14) + Number(this.rowexport_type_14[ax].rement)
      this.summary_plug12_t14 = Number(this.summary_plug12_t14) + Number(this.sum_plug12_t14)
      this.summary_widthloss_t14 = Number(this.summary_widthloss_t14) + Number(this.sum_per_width_loss_t14)
      this.summary_end1_2_t14 = Number(this.summary_end1_2_t14) + Number(this.sum_end1_2_t14)
      this.sum_totallenthloss_t14_first = 0
      this.sum_totallenthloss_t14_first = Number(this.rowexport_type_14[ax].endless_length_yd) + Number(this.rowexport_type_14[ax].spliceloss)
      +Number(this.rowexport_type_14[ax].totalcutout) + Number(this.rowexport_type_14[ax].rement)
      this.sum_totallenthloss_t14 = Number(this.sum_totallenthloss_t14) + Number(this.sum_totallenthloss_t14_first)

      this.sum_totallenthloss_per_t14_first = 0
      this.sum_totallenthloss_per_t14_first = Number(this.rowexport_type_14[ax].avg_end) + Number(this.rowexport_type_14[ax].splicelossper)
      +Number(this.rowexport_type_14[ax].totalcutoutper) + Number(this.rowexport_type_14[ax].rement_per)
      this.sum_totallenthloss_per_t14 = Number(this.sum_totallenthloss_t14) + Number(this.sum_totallenthloss_per_t14_first)
      this.sum_number_1_end_t14 = 0 
      this.sum_number_1_end_t14 = this.rowexport_type_14[ax].endless_length_yd * 36
      this.sum_number_2_end_t14 = 0
      this.sum_number_2_end_t14 = this.rowexport_type_14[ax].ttl_marker/this.rowexport_type_14[ax].length_ydb
      this.sum_plug1_2_end_t14 = 0
      this.sum_plug1_2_end_t14 = this.sum_number_1_end_t14/this.sum_number_2_end_t14

      this.sum_plug1_2_end_total_t14 = 0
     
      this.sum_plug1_2_end_total_t14 = this.sum_plug1_2_end_t14 * this.rowexport_type_14[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t14 = Number(this.sum_all_plug1_2_end_total_t14) + Number(this.sum_plug1_2_end_total_t14)

       this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t14
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t14
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t14
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t14
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t14
      })
       
      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t14,
        value_table:14
      }) 

      ws.getCell("N"+count14).value = Number(this.sum_plug1_2_end_t14)
      ws.getCell("N"+count14).numFmt = '0.00';


      count14++
         }
         if(this.sum_ttl_marker_t14 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t14
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t14 = this.summary_plug12_t14 / this.sum_ttl_marker_t14
        if(this.sum_plug12_t14 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t14
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t14 =  this.summary_widthloss_t14 / this.sum_ttl_marker_t14
        if(this.sum_per_width_loss_t14 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t14
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t14 =  this.summary_end1_2_t14/ this.sum_ttl_marker_t14 //M
      if(this.sum_end1_2_t14 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t14
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t14 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t14
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t14 = this.sum_endless_length_yd_t14 / this.sum_ttl_marker_t14 *100 //P
      if(this.sum_avg_end_t14 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t14
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t14 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t14
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t14 = this.sum_spliceloss_t14 / this.sum_ttl_marker_t14 *100
      if(this.sum_splicelossper_t14 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t14
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t14 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t14
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t14 = this.sum_totalcutout_t14 / this.sum_ttl_marker_t14 *100
      if(this.sum_totalcutoutper_t14 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t14
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t14 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t14
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t14 = this.sum_rement_t14 / this.sum_ttl_marker_t14 *100
  if(this.sum_rement_per_t14 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t14
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t14
    if(this.sum_totallenthloss_t14 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t14
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert14 = this.sum_totallenthlosst14 / this.sum_ttl_marker_t14
    if(this.sum_totallenthlosspert14 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert14
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

ws.getCell("A"+countx14).value = "Total Table A14"  
      ws.mergeCells("A"+countx14+":"+"I"+countx14)
      
      if(this.sum_ttl_marker_t14 > 0 && isNaN(this.sum_ttl_marker_t14)==false && isFinite(this.sum_ttl_marker_t14)==true){
      ws.getCell("J"+countx14).value = Number(this.sum_ttl_marker_t14)
      ws.getCell("J"+countx14).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx14).value = 0
      ws.getCell("J"+countx14).numFmt = '0.00';
      }

      if(this.sum_plug12_t14 > 0 && isNaN(this.sum_plug12_t14)==false && isFinite(this.sum_plug12_t14)==true){
      ws.getCell("K"+countx14).value = Number(this.sum_plug12_t14)
      ws.getCell("K"+countx14).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx14).value = 0
      ws.getCell("K"+countx14).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t14 > 0 && isNaN(this.sum_per_width_loss_t14)==false && isFinite(this.sum_per_width_loss_t14)==true){
      ws.getCell("L"+countx14).value = this.sum_per_width_loss_t14/100
      ws.getCell("L"+countx14).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx14).value = 0.00/100
      ws.getCell("L"+countx14).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t14 = 0
      this.total_result_plug1_2_t14 = this.sum_all_plug1_2_end_total_t14 / this.sum_ttl_marker_t14

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t14
      })
      if(this.total_result_plug1_2_t14 > 0 && isNaN(this.total_result_plug1_2_t14)==false && isFinite(this.total_result_plug1_2_t14)==true){
      ws.getCell("N"+countx14).value = Number(this.total_result_plug1_2_t14)
      ws.getCell("N"+countx14).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx14).value = 0.00
      ws.getCell("N"+countx14).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_t14 > 0 && isNaN(this.sum_endless_length_yd_t14)==false && isFinite(this.sum_endless_length_yd_t14)==true){
      ws.getCell("O"+countx14).value = Number(this.sum_endless_length_yd_t14)
      ws.getCell("O"+countx14).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx14).value = 0.00
      ws.getCell("O"+countx14).numFmt = '0.00';
      }

      if(this.sum_avg_end_t14 > 0 && isNaN(this.sum_avg_end_t14)==false && isFinite(this.sum_avg_end_t14)==true){
      ws.getCell("P"+countx14).value = this.sum_avg_end_t14/100
      ws.getCell("P"+countx14).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx14).value = 0/100
      ws.getCell("P"+countx14).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t14 > 0 && isNaN(this.sum_spliceloss_t14)==false && isFinite(this.sum_spliceloss_t14)==true){
      ws.getCell("R"+countx14).value = Number(this.sum_spliceloss_t14)
      ws.getCell("R"+countx14).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx14).value = 0.00
      ws.getCell("R"+countx14).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t14 > 0 && isNaN(this.sum_splicelossper_t14)==false && isFinite(this.sum_splicelossper_t14)==true){
      ws.getCell("S"+countx14).value = this.sum_splicelossper_t14/100
      ws.getCell("S"+countx14).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx14).value = 0/100
      ws.getCell("S"+countx14).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t14 > 0 && isNaN(this.sum_totalcutout_t14)==false && isFinite(this.sum_totalcutout_t14)==true){
      ws.getCell("U"+countx14).value = Number(this.sum_totalcutout_t14)
      ws.getCell("U"+countx14).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx14).value = 0.00
      ws.getCell("U"+countx14).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t14 > 0 && isNaN(this.sum_totalcutoutper_t14)==false && isFinite(this.sum_totalcutoutper_t14)==true){
      ws.getCell("V"+countx14).value = this.sum_totalcutoutper_t14/100
      ws.getCell("V"+countx14).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx14).value = 0.00/100
      ws.getCell("V"+countx14).numFmt = '0.00%';
      }

       if(this.sum_rement_t14 > 0 && isNaN(this.sum_rement_t14)==false && isFinite(this.sum_rement_t14)==true){
      ws.getCell("X"+countx14).value = Number(this.sum_rement_t14)
      ws.getCell("X"+countx14).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx14).value = 0.00
      ws.getCell("X"+countx14).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t14 > 0 && isNaN(this.sum_rement_per_t14)==false && isFinite(this.sum_rement_per_t14)==true){
      ws.getCell("Y"+countx14).value =  this.sum_rement_per_t14/100
      ws.getCell("Y"+countx14).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx14).value =  0.00/100
      ws.getCell("Y"+countx14).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t14 > 0 && isNaN(this.sum_totallenthloss_t14)==false && isFinite(this.sum_totallenthloss_t14)==true){
      ws.getCell("AA"+countx14).value = Number(this.sum_totallenthloss_t14)
      ws.getCell("AA"+countx14).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx14).value = 0
      ws.getCell("AA"+countx14).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t14 = this.sum_totallenthloss_t14 / this.sum_ttl_marker_t14 *100
      if(this.last_sum_totallenthloss_per_t14 > 0 && isNaN(this.last_sum_totallenthloss_per_t14)==false && isFinite(this.last_sum_totallenthloss_per_t14)==true){
      ws.getCell("AB"+countx14).value = this.last_sum_totallenthloss_per_t14/100
      ws.getCell("AB"+countx14).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx14).value = 0/100
      ws.getCell("AB"+countx14).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx14).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //A15

      if(this.rowexport_type_15.length == 0){
       
         var count15 = count14

       }else{
      var count15 = countx14+1
       }
      
  
        var countx15 = count15 + 1 +Number(this.rowexport_type_15.length)
      
      this.sum_ttl_marker_t15 = 0
      this.sum_plug12_t15 = 0
      this.sum_per_width_loss_t15 = 0
      this.sum_end1_2_t15 = 0
      this.sum_endless_length_yd_t15= 0
      this.sum_avg_end_t15 = 0
      this.sum_spliceloss_t15 = 0
      this.sum_splicelossper_t15 = 0
      this.sum_totalcutout_t15 = 0
      this.sum_totalcutoutper_t15 = 0
      this.sum_rement_t15 = 0
      this.sum_rement_per_t15 = 0
      this.sum_totallenthloss_t15 =0 
      this.sum_totallenthlossper_t15 = 0
      this.summary_widthloss_t15 = 0
      this.summary_plug12_t15 = 0
      this.summary_end1_2_t15 = 0

      
     
      this.sum_all_plug1_2_end_total_t15 = 0
      if(this.rowexport_type_15.length < 1){
        count_row_grand++
      }else{
          for (var ax = 0; ax<this.rowexport_type_15.length; ax++){ //length +120
      ws.getCell("A"+count15).value = this.rowexport_type_15[ax].mu_date

      ws.getCell("B"+count15).value = this.rowexport_type_15[ax].gm_table

      ws.getCell("C"+count15).value = Number(this.rowexport_type_15[ax].gm_no_short)
      ws.getCell("C"+count15).numFmt = '0.00';

      ws.getCell("D"+count15).value = this.rowexport_type_15[ax].customer

      ws.getCell("E"+count15).value = this.rowexport_type_15[ax].gm_no

      ws.getCell("F"+count15).value = this.rowexport_type_15[ax].fabric_type

      ws.getCell("G"+count15).value = Number(this.rowexport_type_15[ax].obs_width)
      ws.getCell("G"+count15).numFmt = '0.00';

      ws.getCell("H"+count15).value = Number(this.rowexport_type_15[ax].width_inch)
      ws.getCell("H"+count15).numFmt = '0.00';


      ws.getCell("I"+count15).value = Number(this.rowexport_type_15[ax].length_ydb)
      ws.getCell("I"+count15).numFmt = '0.00';

      ws.getCell("J"+count15).value = Number(this.rowexport_type_15[ax].ttl_marker)
      ws.getCell("J"+count15).numFmt = '0.00';

      ws.getCell("K"+count15).value = Number(this.rowexport_type_15[ax].plug12)
      ws.getCell("K"+count15).numFmt = '0.00';

      ws.getCell("L"+count15).value = this.rowexport_type_15[ax].widthloss/100
      ws.getCell("L"+count15).numFmt = '0.00%';

      ws.getCell("M"+count15).value = this.rowexport_type_15[ax].widthloss_code
    
      ws.getCell("O"+count15).value = Number(this.rowexport_type_15[ax].endless_length_yd)
      ws.getCell("O"+count15).numFmt = '0.00';

      ws.getCell("P"+count15).value = this.rowexport_type_15[ax].avg_end/100
      ws.getCell("P"+count15).numFmt = '0.00%';

      ws.getCell("Q"+count15).value = this.rowexport_type_15[ax].avg_end_code
      

      ws.getCell("R"+count15).value = Number(this.rowexport_type_15[ax].spliceloss)
      ws.getCell("R"+count15).numFmt = '0.00';

      ws.getCell("S"+count15).value = this.rowexport_type_15[ax].splicelossper/100
      ws.getCell("S"+count15).numFmt = '0.00%';

      ws.getCell("T"+count15).value = this.rowexport_type_15[ax].per_splice_loss_code
   

      ws.getCell("U"+count15).value = Number(this.rowexport_type_15[ax].totalcutout)
      ws.getCell("U"+count15).numFmt = '0.00';

      ws.getCell("V"+count15).value = this.rowexport_type_15[ax].totalcutoutper/100
      ws.getCell("V"+count15).numFmt = '0.00%';

      ws.getCell("W"+count15).value = this.rowexport_type_15[ax].total_cut_out_per_code

      ws.getCell("X"+count15).value = Number(this.rowexport_type_15[ax].rement)
      ws.getCell("X"+count15).numFmt = '0.00';

      ws.getCell("Y"+count15).value = this.rowexport_type_15[ax].rement_per/100
      ws.getCell("Y"+count15).numFmt = '0.00%';

      ws.getCell("Z"+count15).value = this.rowexport_type_15[ax].remnants_loss_code

      ws.getCell("AA"+count15).value = Number(this.rowexport_type_15[ax].totallenthloss)
      ws.getCell("AA"+count15).numFmt = '0.00';

      ws.getCell("AB"+count15).value = this.rowexport_type_15[ax].totallenthlossper/100
      ws.getCell("AB"+count15).numFmt = '0.00%';
      
      this.sum_plug12_t15 = 0
      this.sum_ttl_marker_t15 = Number(this.sum_ttl_marker_t15) + Number(this.rowexport_type_15[ax].ttl_marker)
      this.sum_plug12_t15 = this.rowexport_type_15[ax].plug12 * this.rowexport_type_15[ax].ttl_marker
      this.sum_per_width_loss_t15 = this.rowexport_type_15[ax].widthloss * this.rowexport_type_15[ax].ttl_marker
      this.sum_end1_2_t15 = this.rowexport_type_15[ax].end1_2 * this.rowexport_type_15[ax].ttl_marker
      this.sum_endless_length_yd_t15 = Number(this.sum_endless_length_yd_t15) + Number(this.rowexport_type_15[ax].endless_length_yd)
      this.sum_spliceloss_t15 = Number(this.sum_spliceloss_t15) + Number(this.rowexport_type_15[ax].spliceloss)
      this.sum_totalcutout_t15 = Number(this.sum_totalcutout_t15) + Number(this.rowexport_type_15[ax].totalcutout)
      this.sum_rement_t15 = Number(this.sum_rement_t15) + Number(this.rowexport_type_15[ax].rement)
      this.summary_plug12_t15 = Number(this.summary_plug12_t15) + Number(this.sum_plug12_t15)
      this.summary_widthloss_t15 = Number(this.summary_widthloss_t15) + Number(this.sum_per_width_loss_t15)
      this.summary_end1_2_t15 = Number(this.summary_end1_2_t15) + Number(this.sum_end1_2_t15)
      this.sum_totallenthloss_t15_first = 0
      this.sum_totallenthloss_t15_first = Number(this.rowexport_type_15[ax].endless_length_yd) + Number(this.rowexport_type_15[ax].spliceloss)
      +Number(this.rowexport_type_15[ax].totalcutout) + Number(this.rowexport_type_15[ax].rement)
      this.sum_totallenthloss_t15 = Number(this.sum_totallenthloss_t15) + Number(this.sum_totallenthloss_t15_first)

      this.sum_totallenthloss_per_t15_first = 0
      this.sum_totallenthloss_per_t15_first = Number(this.rowexport_type_15[ax].avg_end) + Number(this.rowexport_type_15[ax].splicelossper)
      +Number(this.rowexport_type_15[ax].totalcutoutper) + Number(this.rowexport_type_15[ax].rement_per)
      this.sum_totallenthloss_per_t15 = Number(this.sum_totallenthloss_t15) + Number(this.sum_totallenthloss_per_t15_first)
      this.sum_number_1_end_t15 = 0 
      this.sum_number_1_end_t15 = this.rowexport_type_15[ax].endless_length_yd * 36
      this.sum_number_2_end_t15 = 0
      this.sum_number_2_end_t15 = this.rowexport_type_15[ax].ttl_marker/this.rowexport_type_15[ax].length_ydb
      this.sum_plug1_2_end_t15 = 0
      this.sum_plug1_2_end_t15 = this.sum_number_1_end_t15/this.sum_number_2_end_t15

      this.sum_plug1_2_end_total_t15 = 0
     
      this.sum_plug1_2_end_total_t15 = this.sum_plug1_2_end_t15 * this.rowexport_type_15[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t15 = Number(this.sum_all_plug1_2_end_total_t15) + Number(this.sum_plug1_2_end_total_t15)

       this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t15
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t15
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t15
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t15
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t15
      })
       
      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t15,
        value_table:15
      }) 

      ws.getCell("N"+count15).value = Number(this.sum_plug1_2_end_t15)
      ws.getCell("N"+count15).numFmt = '0.00';


      count15++
         }
         if(this.sum_ttl_marker_t15 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t15
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t15 = this.summary_plug12_t15 / this.sum_ttl_marker_t15
        if(this.sum_plug12_t15 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t15
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t15 =  this.summary_widthloss_t15 / this.sum_ttl_marker_t15
        if(this.sum_per_width_loss_t15 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t15
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t15 =  this.summary_end1_2_t15/ this.sum_ttl_marker_t15 //M
      if(this.sum_end1_2_t15 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t15
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t15 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t15
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t15 = this.sum_endless_length_yd_t15 / this.sum_ttl_marker_t15 *100 //P
      if(this.sum_avg_end_t15 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t15
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t15 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t15
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t15 = this.sum_spliceloss_t15 / this.sum_ttl_marker_t15 *100
      if(this.sum_splicelossper_t15 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t15
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t15 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t15
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t15 = this.sum_totalcutout_t15 / this.sum_ttl_marker_t15 *100
      if(this.sum_totalcutoutper_t15 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t15
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t15 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t15
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t15 = this.sum_rement_t15 / this.sum_ttl_marker_t15 *100
  if(this.sum_rement_per_t15 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t15
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t15
    if(this.sum_totallenthloss_t15 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t15
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert15 = this.sum_totallenthlosst15 / this.sum_ttl_marker_t15
    if(this.sum_totallenthlosspert15 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert15
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

ws.getCell("A"+countx15).value = "Total Table A15"  
      ws.mergeCells("A"+countx15+":"+"I"+countx15)
      
      if(this.sum_ttl_marker_t15 > 0 && isNaN(this.sum_ttl_marker_t15)==false && isFinite(this.sum_ttl_marker_t15)==true){
      ws.getCell("J"+countx15).value = Number(this.sum_ttl_marker_t15)
      ws.getCell("J"+countx15).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx15).value = 0
      ws.getCell("J"+countx15).numFmt = '0.00';
      }

      if(this.sum_plug12_t15 > 0 && isNaN(this.sum_plug12_t15)==false && isFinite(this.sum_plug12_t15)==true){
      ws.getCell("K"+countx15).value = Number(this.sum_plug12_t15)
      ws.getCell("K"+countx15).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx15).value = 0
      ws.getCell("K"+countx15).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t15 > 0 && isNaN(this.sum_per_width_loss_t15)==false && isFinite(this.sum_per_width_loss_t15)==true){
      ws.getCell("L"+countx15).value = this.sum_per_width_loss_t15/100
      ws.getCell("L"+countx15).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx15).value = 0.00/100
      ws.getCell("L"+countx15).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t15 = 0
      this.total_result_plug1_2_t15 = this.sum_all_plug1_2_end_total_t15 / this.sum_ttl_marker_t15

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t15
      })
      if(this.total_result_plug1_2_t15 > 0 && isNaN(this.total_result_plug1_2_t15)==false && isFinite(this.total_result_plug1_2_t15)==true){
      ws.getCell("N"+countx15).value = Number(this.total_result_plug1_2_t15)
      ws.getCell("N"+countx15).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx15).value = 0.00
      ws.getCell("N"+countx15).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_t15 > 0 && isNaN(this.sum_endless_length_yd_t15)==false && isFinite(this.sum_endless_length_yd_t15)==true){
      ws.getCell("O"+countx15).value = Number(this.sum_endless_length_yd_t15)
      ws.getCell("O"+countx15).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx15).value = 0.00
      ws.getCell("O"+countx15).numFmt = '0.00';
      }

      if(this.sum_avg_end_t15 > 0 && isNaN(this.sum_avg_end_t15)==false && isFinite(this.sum_avg_end_t15)==true){
      ws.getCell("P"+countx15).value = this.sum_avg_end_t15/100
      ws.getCell("P"+countx15).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx15).value = 0/100
      ws.getCell("P"+countx15).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t15 > 0 && isNaN(this.sum_spliceloss_t15)==false && isFinite(this.sum_spliceloss_t15)==true){
      ws.getCell("R"+countx15).value = Number(this.sum_spliceloss_t15)
      ws.getCell("R"+countx15).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx15).value = 0.00
      ws.getCell("R"+countx15).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t15 > 0 && isNaN(this.sum_splicelossper_t15)==false && isFinite(this.sum_splicelossper_t15)==true){
      ws.getCell("S"+countx15).value = this.sum_splicelossper_t15/100
      ws.getCell("S"+countx15).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx15).value = 0/100
      ws.getCell("S"+countx15).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t15 > 0 && isNaN(this.sum_totalcutout_t15)==false && isFinite(this.sum_totalcutout_t15)==true){
      ws.getCell("U"+countx15).value = Number(this.sum_totalcutout_t15)
      ws.getCell("U"+countx15).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx15).value = 0.00
      ws.getCell("U"+countx15).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t15 > 0 && isNaN(this.sum_totalcutoutper_t15)==false && isFinite(this.sum_totalcutoutper_t15)==true){
      ws.getCell("V"+countx15).value = this.sum_totalcutoutper_t15/100
      ws.getCell("V"+countx15).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx15).value = 0.00/100
      ws.getCell("V"+countx15).numFmt = '0.00%';
      }

       if(this.sum_rement_t15 > 0 && isNaN(this.sum_rement_t15)==false && isFinite(this.sum_rement_t15)==true){
      ws.getCell("X"+countx15).value = Number(this.sum_rement_t15)
      ws.getCell("X"+countx15).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx15).value = 0.00
      ws.getCell("X"+countx15).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t15 > 0 && isNaN(this.sum_rement_per_t15)==false && isFinite(this.sum_rement_per_t15)==true){
      ws.getCell("Y"+countx15).value =  this.sum_rement_per_t15/100
      ws.getCell("Y"+countx15).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx15).value =  0.00/100
      ws.getCell("Y"+countx15).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t15 > 0 && isNaN(this.sum_totallenthloss_t15)==false && isFinite(this.sum_totallenthloss_t15)==true){
      ws.getCell("AA"+countx15).value = Number(this.sum_totallenthloss_t15)
      ws.getCell("AA"+countx15).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx15).value = 0
      ws.getCell("AA"+countx15).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t15 = this.sum_totallenthloss_t15 / this.sum_ttl_marker_t15 *100
      if(this.last_sum_totallenthloss_per_t15 > 0 && isNaN(this.last_sum_totallenthloss_per_t15)==false && isFinite(this.last_sum_totallenthloss_per_t15)==true){
      ws.getCell("AB"+countx15).value = this.last_sum_totallenthloss_per_t15/100
      ws.getCell("AB"+countx15).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx15).value = 0/100
      ws.getCell("AB"+countx15).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx15).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }
      //A16

      if(this.rowexport_type_15.length == 0){
       
         var count16 = count15

       }else{
       var count16 = countx15+1
       }
     
  
        var countx16 = count16 + 1 +Number(this.rowexport_type_16.length)
      
      this.sum_ttl_marker_t16 = 0
      this.sum_plug12_t16 = 0
      this.sum_per_width_loss_t16 = 0
      this.sum_end1_2_t16 = 0
      this.sum_endless_length_yd_t16= 0
      this.sum_avg_end_t16 = 0
      this.sum_spliceloss_t16 = 0
      this.sum_splicelossper_t16 = 0
      this.sum_totalcutout_t16 = 0
      this.sum_totalcutoutper_t16 = 0
      this.sum_rement_t16 = 0
      this.sum_rement_per_t16 = 0
      this.sum_totallenthloss_t16 =0 
      this.sum_totallenthlossper_t16 = 0
      this.summary_widthloss_t16 = 0
      this.summary_plug12_t16 = 0
      this.summary_end1_2_t16 = 0

      
     
      this.sum_all_plug1_2_end_total_t16 = 0

      if(this.rowexport_type_16.length < 1){
        console.log(count_row_grand)
        count_row_grand++
        
      }else{
     
          for (var ax = 0; ax<this.rowexport_type_16.length; ax++){ //length +120
      ws.getCell("A"+count16).value = this.rowexport_type_16[ax].mu_date

      ws.getCell("B"+count16).value = this.rowexport_type_16[ax].gm_table

      ws.getCell("C"+count16).value = Number(this.rowexport_type_16[ax].gm_no_short)
      ws.getCell("C"+count16).numFmt = '0.00';

      ws.getCell("D"+count16).value = this.rowexport_type_16[ax].customer

      ws.getCell("E"+count16).value = this.rowexport_type_16[ax].gm_no

      ws.getCell("F"+count16).value = this.rowexport_type_16[ax].fabric_type

      ws.getCell("G"+count16).value = Number(this.rowexport_type_16[ax].obs_width)
      ws.getCell("G"+count16).numFmt = '0.00';

      ws.getCell("H"+count16).value = Number(this.rowexport_type_16[ax].width_inch)
      ws.getCell("H"+count16).numFmt = '0.00';


      ws.getCell("I"+count16).value = Number(this.rowexport_type_16[ax].length_ydb)
      ws.getCell("I"+count16).numFmt = '0.00';

      ws.getCell("J"+count16).value = Number(this.rowexport_type_16[ax].ttl_marker)
      ws.getCell("J"+count16).numFmt = '0.00';

      ws.getCell("K"+count16).value = Number(this.rowexport_type_16[ax].plug12)
      ws.getCell("K"+count16).numFmt = '0.00';

      ws.getCell("L"+count16).value = this.rowexport_type_16[ax].widthloss/100
      ws.getCell("L"+count16).numFmt = '0.00%';

      ws.getCell("M"+count16).value = this.rowexport_type_16[ax].widthloss_code
    
      ws.getCell("O"+count16).value = Number(this.rowexport_type_16[ax].endless_length_yd)
      ws.getCell("O"+count16).numFmt = '0.00';

      ws.getCell("P"+count16).value = this.rowexport_type_16[ax].avg_end/100
      ws.getCell("P"+count16).numFmt = '0.00%';

      ws.getCell("Q"+count16).value = this.rowexport_type_16[ax].avg_end_code
      

      ws.getCell("R"+count16).value = Number(this.rowexport_type_16[ax].spliceloss)
      ws.getCell("R"+count16).numFmt = '0.00';

      ws.getCell("S"+count16).value = this.rowexport_type_16[ax].splicelossper/100
      ws.getCell("S"+count16).numFmt = '0.00%';

      ws.getCell("T"+count16).value = this.rowexport_type_16[ax].per_splice_loss_code
   

      ws.getCell("U"+count16).value = Number(this.rowexport_type_16[ax].totalcutout)
      ws.getCell("U"+count16).numFmt = '0.00';

      ws.getCell("V"+count16).value = this.rowexport_type_16[ax].totalcutoutper/100
      ws.getCell("V"+count16).numFmt = '0.00%';

      ws.getCell("W"+count16).value = this.rowexport_type_16[ax].total_cut_out_per_code

      ws.getCell("X"+count16).value = Number(this.rowexport_type_16[ax].rement)
      ws.getCell("X"+count16).numFmt = '0.00';

      ws.getCell("Y"+count16).value = this.rowexport_type_16[ax].rement_per/100
      ws.getCell("Y"+count16).numFmt = '0.00%';

      ws.getCell("Z"+count16).value = this.rowexport_type_16[ax].remnants_loss_code

      ws.getCell("AA"+count16).value = Number(this.rowexport_type_16[ax].totallenthloss)
      ws.getCell("AA"+count16).numFmt = '0.00';

      ws.getCell("AB"+count16).value = this.rowexport_type_16[ax].totallenthlossper/100
      ws.getCell("AB"+count16).numFmt = '0.00%';
      
      this.sum_plug12_t16 = 0
      this.sum_ttl_marker_t16 = Number(this.sum_ttl_marker_t16) + Number(this.rowexport_type_16[ax].ttl_marker)
      this.sum_plug12_t16 = this.rowexport_type_16[ax].plug12 * this.rowexport_type_16[ax].ttl_marker
      this.sum_per_width_loss_t16 = this.rowexport_type_16[ax].widthloss * this.rowexport_type_16[ax].ttl_marker
      this.sum_end1_2_t16 = this.rowexport_type_16[ax].end1_2 * this.rowexport_type_16[ax].ttl_marker
      this.sum_endless_length_yd_t16 = Number(this.sum_endless_length_yd_t16) + Number(this.rowexport_type_16[ax].endless_length_yd)
      this.sum_spliceloss_t16 = Number(this.sum_spliceloss_t16) + Number(this.rowexport_type_16[ax].spliceloss)
      this.sum_totalcutout_t16 = Number(this.sum_totalcutout_t16) + Number(this.rowexport_type_16[ax].totalcutout)
      this.sum_rement_t16 = Number(this.sum_rement_t16) + Number(this.rowexport_type_16[ax].rement)
      this.summary_plug12_t16 = Number(this.summary_plug12_t16) + Number(this.sum_plug12_t16)
      this.summary_widthloss_t16 = Number(this.summary_widthloss_t16) + Number(this.sum_per_width_loss_t16)
      this.summary_end1_2_t16 = Number(this.summary_end1_2_t16) + Number(this.sum_end1_2_t16)
      this.sum_totallenthloss_t16_first = 0
      this.sum_totallenthloss_t16_first = Number(this.rowexport_type_16[ax].endless_length_yd) + Number(this.rowexport_type_16[ax].spliceloss)
      +Number(this.rowexport_type_16[ax].totalcutout) + Number(this.rowexport_type_16[ax].rement)
      this.sum_totallenthloss_t16 = Number(this.sum_totallenthloss_t16) + Number(this.sum_totallenthloss_t16_first)

      this.sum_totallenthloss_per_t16_first = 0
      this.sum_totallenthloss_per_t16_first = Number(this.rowexport_type_16[ax].avg_end) + Number(this.rowexport_type_16[ax].splicelossper)
      +Number(this.rowexport_type_16[ax].totalcutoutper) + Number(this.rowexport_type_16[ax].rement_per)
      this.sum_totallenthloss_per_t16 = Number(this.sum_totallenthloss_t16) + Number(this.sum_totallenthloss_per_t16_first)
      this.sum_number_1_end_t16 = 0 
      this.sum_number_1_end_t16 = this.rowexport_type_16[ax].endless_length_yd * 36
      this.sum_number_2_end_t16 = 0
      this.sum_number_2_end_t16 = this.rowexport_type_16[ax].ttl_marker/this.rowexport_type_16[ax].length_ydb
      this.sum_plug1_2_end_t16 = 0
      this.sum_plug1_2_end_t16 = this.sum_number_1_end_t16/this.sum_number_2_end_t16

      this.sum_plug1_2_end_total_t16 = 0
     
      this.sum_plug1_2_end_total_t16 = this.sum_plug1_2_end_t16 * this.rowexport_type_16[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t16 = Number(this.sum_all_plug1_2_end_total_t16) + Number(this.sum_plug1_2_end_total_t16)

       this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t16
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t16
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t16
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t16
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t16
      })
       
      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t16,
        value_table:16
      }) 

      ws.getCell("N"+count16).value = Number(this.sum_plug1_2_end_t16)
      ws.getCell("N"+count16).numFmt = '0.00';


      count16++
         }
         if(this.sum_ttl_marker_t16 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t16
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t16 = this.summary_plug12_t16 / this.sum_ttl_marker_t16
        if(this.sum_plug12_t16 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t16
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t16 =  this.summary_widthloss_t16 / this.sum_ttl_marker_t16
        if(this.sum_per_width_loss_t16 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t16
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t16 =  this.summary_end1_2_t16/ this.sum_ttl_marker_t16 //M
      if(this.sum_end1_2_t16 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t16
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t16 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t16
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t16 = this.sum_endless_length_yd_t16 / this.sum_ttl_marker_t16 *100 //P
      if(this.sum_avg_end_t16 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t16
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t16 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t16
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t16 = this.sum_spliceloss_t16 / this.sum_ttl_marker_t16 *100
      if(this.sum_splicelossper_t16 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t16
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t16 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t16
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t16 = this.sum_totalcutout_t16 / this.sum_ttl_marker_t16 *100
      if(this.sum_totalcutoutper_t16 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t16
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t16 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t16
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t16 = this.sum_rement_t16 / this.sum_ttl_marker_t16 *100
  if(this.sum_rement_per_t16 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t16
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t16
    if(this.sum_totallenthloss_t16 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t16
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert16 = this.sum_totallenthlosst16 / this.sum_ttl_marker_t16
    if(this.sum_totallenthlosspert16 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert16
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

ws.getCell("A"+countx16).value = "Total Table A16"  
      ws.mergeCells("A"+countx16+":"+"I"+countx16)
      
      if(this.sum_ttl_marker_t16 > 0 && isNaN(this.sum_ttl_marker_t16)==false && isFinite(this.sum_ttl_marker_t16)==true){
      ws.getCell("J"+countx16).value = Number(this.sum_ttl_marker_t16)
      ws.getCell("J"+countx16).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx16).value = 0
      ws.getCell("J"+countx16).numFmt = '0.00';
      }

      if(this.sum_plug12_t16 > 0 && isNaN(this.sum_plug12_t16)==false && isFinite(this.sum_plug12_t16)==true){
      ws.getCell("K"+countx16).value = Number(this.sum_plug12_t16)
      ws.getCell("K"+countx16).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx16).value = 0
      ws.getCell("K"+countx16).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t16 > 0 && isNaN(this.sum_per_width_loss_t16)==false && isFinite(this.sum_per_width_loss_t16)==true){
      ws.getCell("L"+countx16).value = this.sum_per_width_loss_t16/100
      ws.getCell("L"+countx16).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx16).value = 0.00/100
      ws.getCell("L"+countx16).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t16 = 0
      this.total_result_plug1_2_t16 = this.sum_all_plug1_2_end_total_t16 / this.sum_ttl_marker_t16

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t16
      })
      if(this.total_result_plug1_2_t16 > 0 && isNaN(this.total_result_plug1_2_t16)==false && isFinite(this.total_result_plug1_2_t16)==true){
      ws.getCell("N"+countx16).value = Number(this.total_result_plug1_2_t16)
      ws.getCell("N"+countx16).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx16).value = 0.00
      ws.getCell("N"+countx16).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_t16 > 0 && isNaN(this.sum_endless_length_yd_t16)==false && isFinite(this.sum_endless_length_yd_t16)==true){
      ws.getCell("O"+countx16).value = Number(this.sum_endless_length_yd_t16)
      ws.getCell("O"+countx16).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx16).value = 0.00
      ws.getCell("O"+countx16).numFmt = '0.00';
      }

      if(this.sum_avg_end_t16 > 0 && isNaN(this.sum_avg_end_t16)==false && isFinite(this.sum_avg_end_t16)==true){
      ws.getCell("P"+countx16).value = this.sum_avg_end_t16/100
      ws.getCell("P"+countx16).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx16).value = 0/100
      ws.getCell("P"+countx16).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t16 > 0 && isNaN(this.sum_spliceloss_t16)==false && isFinite(this.sum_spliceloss_t16)==true){
      ws.getCell("R"+countx16).value = Number(this.sum_spliceloss_t16)
      ws.getCell("R"+countx16).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx16).value = 0.00
      ws.getCell("R"+countx16).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t16 > 0 && isNaN(this.sum_splicelossper_t16)==false && isFinite(this.sum_splicelossper_t16)==true){
      ws.getCell("S"+countx16).value = this.sum_splicelossper_t16/100
      ws.getCell("S"+countx16).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx16).value = 0/100
      ws.getCell("S"+countx16).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t16 > 0 && isNaN(this.sum_totalcutout_t16)==false && isFinite(this.sum_totalcutout_t16)==true){
      ws.getCell("U"+countx16).value = Number(this.sum_totalcutout_t16)
      ws.getCell("U"+countx16).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx16).value = 0.00
      ws.getCell("U"+countx16).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t16 > 0 && isNaN(this.sum_totalcutoutper_t16)==false && isFinite(this.sum_totalcutoutper_t16)==true){
      ws.getCell("V"+countx16).value = this.sum_totalcutoutper_t16/100
      ws.getCell("V"+countx16).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx16).value = 0.00/100
      ws.getCell("V"+countx16).numFmt = '0.00%';
      }

       if(this.sum_rement_t16 > 0 && isNaN(this.sum_rement_t16)==false && isFinite(this.sum_rement_t16)==true){
      ws.getCell("X"+countx16).value = Number(this.sum_rement_t16)
      ws.getCell("X"+countx16).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx16).value = 0.00
      ws.getCell("X"+countx16).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t16 > 0 && isNaN(this.sum_rement_per_t16)==false && isFinite(this.sum_rement_per_t16)==true){
      ws.getCell("Y"+countx16).value =  this.sum_rement_per_t16/100
      ws.getCell("Y"+countx16).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx16).value =  0.00/100
      ws.getCell("Y"+countx16).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t16 > 0 && isNaN(this.sum_totallenthloss_t16)==false && isFinite(this.sum_totallenthloss_t16)==true){
      ws.getCell("AA"+countx16).value = Number(this.sum_totallenthloss_t16)
      ws.getCell("AA"+countx16).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx16).value = 0
      ws.getCell("AA"+countx16).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t16 = this.sum_totallenthloss_t16 / this.sum_ttl_marker_t16 *100
      if(this.last_sum_totallenthloss_per_t16 > 0 && isNaN(this.last_sum_totallenthloss_per_t16)==false && isFinite(this.last_sum_totallenthloss_per_t16)==true){
      ws.getCell("AB"+countx16).value = this.last_sum_totallenthloss_per_t16/100
      ws.getCell("AB"+countx16).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx16).value = 0/100
      ws.getCell("AB"+countx16).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx16).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //A17
      if(this.rowexport_type_16.length == 0){
       
         var count17 = count16

       }else{
       var count17 = countx16+1
       }
     
      
  
        var countx17 = count17 + 1 +Number(this.rowexport_type_17.length)
      
      this.sum_ttl_marker_t17 = 0
      this.sum_plug12_t17 = 0
      this.sum_per_width_loss_t17 = 0
      this.sum_end1_2_t17 = 0
      this.sum_endless_length_yd_t17= 0
      this.sum_avg_end_t17 = 0
      this.sum_spliceloss_t17 = 0
      this.sum_splicelossper_t17 = 0
      this.sum_totalcutout_t17 = 0
      this.sum_totalcutoutper_t17 = 0
      this.sum_rement_t17 = 0
      this.sum_rement_per_t17 = 0
      this.sum_totallenthloss_t17 =0 
      this.sum_totallenthlossper_t17 = 0
      this.summary_widthloss_t17 = 0
      this.summary_plug12_t17 = 0
      this.summary_end1_2_t17 = 0

      
     
      this.sum_all_plug1_2_end_total_t17 = 0

       if(this.rowexport_type_17.length < 1){
      
      count_row_grand++
      
      }else{
        
     
          for (var ax = 0; ax<this.rowexport_type_17.length; ax++){ //length +120
      ws.getCell("A"+count17).value = this.rowexport_type_17[ax].mu_date

      ws.getCell("B"+count17).value = this.rowexport_type_17[ax].gm_table

      ws.getCell("C"+count17).value = Number(this.rowexport_type_17[ax].gm_no_short)
      ws.getCell("C"+count17).numFmt = '0.00';

      ws.getCell("D"+count17).value = this.rowexport_type_17[ax].customer

      ws.getCell("E"+count17).value = this.rowexport_type_17[ax].gm_no

      ws.getCell("F"+count17).value = this.rowexport_type_17[ax].fabric_type

      ws.getCell("G"+count17).value = Number(this.rowexport_type_17[ax].obs_width)
      ws.getCell("G"+count17).numFmt = '0.00';

      ws.getCell("H"+count17).value = Number(this.rowexport_type_17[ax].width_inch)
      ws.getCell("H"+count17).numFmt = '0.00';


      ws.getCell("I"+count17).value = Number(this.rowexport_type_17[ax].length_ydb)
      ws.getCell("I"+count17).numFmt = '0.00';

      ws.getCell("J"+count17).value = Number(this.rowexport_type_17[ax].ttl_marker)
      ws.getCell("J"+count17).numFmt = '0.00';

      ws.getCell("K"+count17).value = Number(this.rowexport_type_17[ax].plug12)
      ws.getCell("K"+count17).numFmt = '0.00';

      ws.getCell("L"+count17).value = this.rowexport_type_17[ax].widthloss/100
      ws.getCell("L"+count17).numFmt = '0.00%';

      ws.getCell("M"+count17).value = this.rowexport_type_17[ax].widthloss_code
    
      ws.getCell("O"+count17).value = Number(this.rowexport_type_17[ax].endless_length_yd)
      ws.getCell("O"+count17).numFmt = '0.00';

      ws.getCell("P"+count17).value = this.rowexport_type_17[ax].avg_end/100
      ws.getCell("P"+count17).numFmt = '0.00%';

      ws.getCell("Q"+count17).value = this.rowexport_type_17[ax].avg_end_code
      

      ws.getCell("R"+count17).value = Number(this.rowexport_type_17[ax].spliceloss)
      ws.getCell("R"+count17).numFmt = '0.00';

      ws.getCell("S"+count17).value = this.rowexport_type_17[ax].splicelossper/100
      ws.getCell("S"+count17).numFmt = '0.00%';

      ws.getCell("T"+count17).value = this.rowexport_type_17[ax].per_splice_loss_code
   

      ws.getCell("U"+count17).value = Number(this.rowexport_type_17[ax].totalcutout)
      ws.getCell("U"+count17).numFmt = '0.00';

      ws.getCell("V"+count17).value = this.rowexport_type_17[ax].totalcutoutper/100
      ws.getCell("V"+count17).numFmt = '0.00%';

      ws.getCell("W"+count17).value = this.rowexport_type_17[ax].total_cut_out_per_code

      ws.getCell("X"+count17).value = Number(this.rowexport_type_17[ax].rement)
      ws.getCell("X"+count17).numFmt = '0.00';

      ws.getCell("Y"+count17).value = this.rowexport_type_17[ax].rement_per/100
      ws.getCell("Y"+count17).numFmt = '0.00%';

      ws.getCell("Z"+count17).value = this.rowexport_type_17[ax].remnants_loss_code

      ws.getCell("AA"+count17).value = Number(this.rowexport_type_17[ax].totallenthloss)
      ws.getCell("AA"+count17).numFmt = '0.00';

      ws.getCell("AB"+count17).value = this.rowexport_type_17[ax].totallenthlossper/100
      ws.getCell("AB"+count17).numFmt = '0.00%';
      
      this.sum_plug12_t17 = 0
      this.sum_ttl_marker_t17 = Number(this.sum_ttl_marker_t17) + Number(this.rowexport_type_17[ax].ttl_marker)
      this.sum_plug12_t17 = this.rowexport_type_17[ax].plug12 * this.rowexport_type_17[ax].ttl_marker
      this.sum_per_width_loss_t17 = this.rowexport_type_17[ax].widthloss * this.rowexport_type_17[ax].ttl_marker
      this.sum_end1_2_t17 = this.rowexport_type_17[ax].end1_2 * this.rowexport_type_17[ax].ttl_marker
      this.sum_endless_length_yd_t17 = Number(this.sum_endless_length_yd_t17) + Number(this.rowexport_type_17[ax].endless_length_yd)
      this.sum_spliceloss_t17 = Number(this.sum_spliceloss_t17) + Number(this.rowexport_type_17[ax].spliceloss)
      this.sum_totalcutout_t17 = Number(this.sum_totalcutout_t17) + Number(this.rowexport_type_17[ax].totalcutout)
      this.sum_rement_t17 = Number(this.sum_rement_t17) + Number(this.rowexport_type_17[ax].rement)
      this.summary_plug12_t17 = Number(this.summary_plug12_t17) + Number(this.sum_plug12_t17)
      this.summary_widthloss_t17 = Number(this.summary_widthloss_t17) + Number(this.sum_per_width_loss_t17)
      this.summary_end1_2_t17 = Number(this.summary_end1_2_t17) + Number(this.sum_end1_2_t17)
      this.sum_totallenthloss_t17_first = 0
      this.sum_totallenthloss_t17_first = Number(this.rowexport_type_17[ax].endless_length_yd) + Number(this.rowexport_type_17[ax].spliceloss)
      +Number(this.rowexport_type_17[ax].totalcutout) + Number(this.rowexport_type_17[ax].rement)
      this.sum_totallenthloss_t17 = Number(this.sum_totallenthloss_t17) + Number(this.sum_totallenthloss_t17_first)

      this.sum_totallenthloss_per_t17_first = 0
      this.sum_totallenthloss_per_t17_first = Number(this.rowexport_type_17[ax].avg_end) + Number(this.rowexport_type_17[ax].splicelossper)
      +Number(this.rowexport_type_17[ax].totalcutoutper) + Number(this.rowexport_type_17[ax].rement_per)
      this.sum_totallenthloss_per_t17 = Number(this.sum_totallenthloss_t17) + Number(this.sum_totallenthloss_per_t17_first)
      this.sum_number_1_end_t17 = 0 
      this.sum_number_1_end_t17 = this.rowexport_type_17[ax].endless_length_yd * 36
      this.sum_number_2_end_t17 = 0
      this.sum_number_2_end_t17 = this.rowexport_type_17[ax].ttl_marker/this.rowexport_type_17[ax].length_ydb
      this.sum_plug1_2_end_t17 = 0
      this.sum_plug1_2_end_t17 = this.sum_number_1_end_t17/this.sum_number_2_end_t17

      this.sum_plug1_2_end_total_t17 = 0
     
      this.sum_plug1_2_end_total_t17 = this.sum_plug1_2_end_t17 * this.rowexport_type_17[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t17 = Number(this.sum_all_plug1_2_end_total_t17) + Number(this.sum_plug1_2_end_total_t17)

       this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t17
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t17
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t17
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t17
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t17
      })
       
      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t17,
        value_table:17
      }) 

      ws.getCell("N"+count17).value = Number(this.sum_plug1_2_end_t17)
      ws.getCell("N"+count17).numFmt = '0.00';


      count17++
         }
         if(this.sum_ttl_marker_t17 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t17
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t17 = this.summary_plug12_t17 / this.sum_ttl_marker_t17
        if(this.sum_plug12_t17 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t17
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t17 =  this.summary_widthloss_t17 / this.sum_ttl_marker_t17
        if(this.sum_per_width_loss_t17 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t17
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t17 =  this.summary_end1_2_t17/ this.sum_ttl_marker_t17 //M
      if(this.sum_end1_2_t17 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t17
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t17 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t17
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t17 = this.sum_endless_length_yd_t17 / this.sum_ttl_marker_t17 *100 //P
      if(this.sum_avg_end_t17 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t17
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t17 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t17
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t17 = this.sum_spliceloss_t17 / this.sum_ttl_marker_t17 *100
      if(this.sum_splicelossper_t17 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t17
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t17 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t17
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t17 = this.sum_totalcutout_t17 / this.sum_ttl_marker_t17 *100
      if(this.sum_totalcutoutper_t17 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t17
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t17 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t17
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t17 = this.sum_rement_t17 / this.sum_ttl_marker_t17 *100
  if(this.sum_rement_per_t17 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t17
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t17
    if(this.sum_totallenthloss_t17 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t17
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert17 = this.sum_totallenthlosst17 / this.sum_ttl_marker_t17
    if(this.sum_totallenthlosspert17 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert17
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

ws.getCell("A"+countx17).value = "Total Table A17"  
      ws.mergeCells("A"+countx17+":"+"I"+countx17)
      
      if(this.sum_ttl_marker_t17 > 0 && isNaN(this.sum_ttl_marker_t17)==false && isFinite(this.sum_ttl_marker_t17)==true){
      ws.getCell("J"+countx17).value = Number(this.sum_ttl_marker_t17)
      ws.getCell("J"+countx17).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx17).value = 0
      ws.getCell("J"+countx17).numFmt = '0.00';
      }

      if(this.sum_plug12_t17 > 0 && isNaN(this.sum_plug12_t17)==false && isFinite(this.sum_plug12_t17)==true){
      ws.getCell("K"+countx17).value = Number(this.sum_plug12_t17)
      ws.getCell("K"+countx17).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx17).value = 0
      ws.getCell("K"+countx17).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t17 > 0 && isNaN(this.sum_per_width_loss_t17)==false && isFinite(this.sum_per_width_loss_t17)==true){
      ws.getCell("L"+countx17).value = this.sum_per_width_loss_t17/100
      ws.getCell("L"+countx17).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx17).value = 0.00/100
      ws.getCell("L"+countx17).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t17 = 0
      this.total_result_plug1_2_t17 = this.sum_all_plug1_2_end_total_t17 / this.sum_ttl_marker_t17

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t17
      })
      if(this.total_result_plug1_2_t17 > 0 && isNaN(this.total_result_plug1_2_t17)==false && isFinite(this.total_result_plug1_2_t17)==true){
      ws.getCell("N"+countx17).value = Number(this.total_result_plug1_2_t17)
      ws.getCell("N"+countx17).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx17).value = 0.00
      ws.getCell("N"+countx17).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_t17 > 0 && isNaN(this.sum_endless_length_yd_t17)==false && isFinite(this.sum_endless_length_yd_t17)==true){
      ws.getCell("O"+countx17).value = Number(this.sum_endless_length_yd_t17)
      ws.getCell("O"+countx17).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx17).value = 0.00
      ws.getCell("O"+countx17).numFmt = '0.00';
      }

      if(this.sum_avg_end_t17 > 0 && isNaN(this.sum_avg_end_t17)==false && isFinite(this.sum_avg_end_t17)==true){
      ws.getCell("P"+countx17).value = this.sum_avg_end_t17/100
      ws.getCell("P"+countx17).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx17).value = 0/100
      ws.getCell("P"+countx17).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t17 > 0 && isNaN(this.sum_spliceloss_t17)==false && isFinite(this.sum_spliceloss_t17)==true){
      ws.getCell("R"+countx17).value = Number(this.sum_spliceloss_t17)
      ws.getCell("R"+countx17).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx17).value = 0.00
      ws.getCell("R"+countx17).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t17 > 0 && isNaN(this.sum_splicelossper_t17)==false && isFinite(this.sum_splicelossper_t17)==true){
      ws.getCell("S"+countx17).value = this.sum_splicelossper_t17/100
      ws.getCell("S"+countx17).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx17).value = 0/100
      ws.getCell("S"+countx17).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t17 > 0 && isNaN(this.sum_totalcutout_t17)==false && isFinite(this.sum_totalcutout_t17)==true){
      ws.getCell("U"+countx17).value = Number(this.sum_totalcutout_t17)
      ws.getCell("U"+countx17).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx17).value = 0.00
      ws.getCell("U"+countx17).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t17 > 0 && isNaN(this.sum_totalcutoutper_t17)==false && isFinite(this.sum_totalcutoutper_t17)==true){
      ws.getCell("V"+countx17).value = this.sum_totalcutoutper_t17/100
      ws.getCell("V"+countx17).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx17).value = 0.00/100
      ws.getCell("V"+countx17).numFmt = '0.00%';
      }

       if(this.sum_rement_t17 > 0 && isNaN(this.sum_rement_t17)==false && isFinite(this.sum_rement_t17)==true){
      ws.getCell("X"+countx17).value = Number(this.sum_rement_t17)
      ws.getCell("X"+countx17).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx17).value = 0.00
      ws.getCell("X"+countx17).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t17 > 0 && isNaN(this.sum_rement_per_t17)==false && isFinite(this.sum_rement_per_t17)==true){
      ws.getCell("Y"+countx17).value =  this.sum_rement_per_t17/100
      ws.getCell("Y"+countx17).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx17).value =  0.00/100
      ws.getCell("Y"+countx17).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t17 > 0 && isNaN(this.sum_totallenthloss_t17)==false && isFinite(this.sum_totallenthloss_t17)==true){
      ws.getCell("AA"+countx17).value = Number(this.sum_totallenthloss_t17)
      ws.getCell("AA"+countx17).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx17).value = 0
      ws.getCell("AA"+countx17).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t17 = this.sum_totallenthloss_t17 / this.sum_ttl_marker_t17 *100
      if(this.last_sum_totallenthloss_per_t17 > 0 && isNaN(this.last_sum_totallenthloss_per_t17)==false && isFinite(this.last_sum_totallenthloss_per_t17)==true){
      ws.getCell("AB"+countx17).value = this.last_sum_totallenthloss_per_t17/100
      ws.getCell("AB"+countx17).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx17).value = 0/100
      ws.getCell("AB"+countx17).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx17).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
       }

      //A18
      if(this.rowexport_type_17.length == 0){
       
         var count18 = count17

       }else{
       var count18 = countx17+1
       }
       
  
        var countx18 = count18 + 1 +Number(this.rowexport_type_18.length)
    
      this.sum_ttl_marker_t18 = 0
      this.sum_plug12_t18 = 0
      this.sum_per_width_loss_t18 = 0
      this.sum_end1_2_t18 = 0
      this.sum_endless_length_yd_t18= 0
      this.sum_avg_end_t18 = 0
      this.sum_spliceloss_t18 = 0
      this.sum_splicelossper_t18 = 0
      this.sum_totalcutout_t18 = 0
      this.sum_totalcutoutper_t18 = 0
      this.sum_rement_t18 = 0
      this.sum_rement_per_t18 = 0
      this.sum_totallenthloss_t18 =0 
      this.sum_totallenthlossper_t18 = 0
      this.summary_widthloss_t18 = 0
      this.summary_plug12_t18 = 0
      this.summary_end1_2_t18 = 0

      
     
      this.sum_all_plug1_2_end_total_t18 = 0

         if(this.rowexport_type_18.length < 1){
      count_row_grand++
      
      
      }else{
        
          for (var ax = 0; ax<this.rowexport_type_18.length; ax++){ //length +120
      ws.getCell("A"+count18).value = this.rowexport_type_18[ax].mu_date

      ws.getCell("B"+count18).value = this.rowexport_type_18[ax].gm_table

      ws.getCell("C"+count18).value = Number(this.rowexport_type_18[ax].gm_no_short)
      ws.getCell("C"+count18).numFmt = '0.00';

      ws.getCell("D"+count18).value = this.rowexport_type_18[ax].customer

      ws.getCell("E"+count18).value = this.rowexport_type_18[ax].gm_no

      ws.getCell("F"+count18).value = this.rowexport_type_18[ax].fabric_type

      ws.getCell("G"+count18).value = Number(this.rowexport_type_18[ax].obs_width)
      ws.getCell("G"+count18).numFmt = '0.00';

      ws.getCell("H"+count18).value = Number(this.rowexport_type_18[ax].width_inch)
      ws.getCell("H"+count18).numFmt = '0.00';


      ws.getCell("I"+count18).value = Number(this.rowexport_type_18[ax].length_ydb)
      ws.getCell("I"+count18).numFmt = '0.00';

      ws.getCell("J"+count18).value = Number(this.rowexport_type_18[ax].ttl_marker)
      ws.getCell("J"+count18).numFmt = '0.00';

      ws.getCell("K"+count18).value = Number(this.rowexport_type_18[ax].plug12)
      ws.getCell("K"+count18).numFmt = '0.00';

      ws.getCell("L"+count18).value = this.rowexport_type_18[ax].widthloss/100
      ws.getCell("L"+count18).numFmt = '0.00%';

      ws.getCell("M"+count18).value = this.rowexport_type_18[ax].widthloss_code
    
      ws.getCell("O"+count18).value = Number(this.rowexport_type_18[ax].endless_length_yd)
      ws.getCell("O"+count18).numFmt = '0.00';

      ws.getCell("P"+count18).value = this.rowexport_type_18[ax].avg_end/100
      ws.getCell("P"+count18).numFmt = '0.00%';

      ws.getCell("Q"+count18).value = this.rowexport_type_18[ax].avg_end_code
      

      ws.getCell("R"+count18).value = Number(this.rowexport_type_18[ax].spliceloss)
      ws.getCell("R"+count18).numFmt = '0.00';

      ws.getCell("S"+count18).value = this.rowexport_type_18[ax].splicelossper/100
      ws.getCell("S"+count18).numFmt = '0.00%';

      ws.getCell("T"+count18).value = this.rowexport_type_18[ax].per_splice_loss_code
   

      ws.getCell("U"+count18).value = Number(this.rowexport_type_18[ax].totalcutout)
      ws.getCell("U"+count18).numFmt = '0.00';

      ws.getCell("V"+count18).value = this.rowexport_type_18[ax].totalcutoutper/100
      ws.getCell("V"+count18).numFmt = '0.00%';

      ws.getCell("W"+count18).value = this.rowexport_type_18[ax].total_cut_out_per_code

      ws.getCell("X"+count18).value = Number(this.rowexport_type_18[ax].rement)
      ws.getCell("X"+count18).numFmt = '0.00';

      ws.getCell("Y"+count18).value = this.rowexport_type_18[ax].rement_per/100
      ws.getCell("Y"+count18).numFmt = '0.00%';

      ws.getCell("Z"+count18).value = this.rowexport_type_18[ax].remnants_loss_code

      ws.getCell("AA"+count18).value = Number(this.rowexport_type_18[ax].totallenthloss)
      ws.getCell("AA"+count18).numFmt = '0.00';

      ws.getCell("AB"+count18).value = this.rowexport_type_18[ax].totallenthlossper/100
      ws.getCell("AB"+count18).numFmt = '0.00%';
      
      this.sum_plug12_t18 = 0
      this.sum_ttl_marker_t18 = Number(this.sum_ttl_marker_t18) + Number(this.rowexport_type_18[ax].ttl_marker)
      this.sum_plug12_t18 = this.rowexport_type_18[ax].plug12 * this.rowexport_type_18[ax].ttl_marker
      this.sum_per_width_loss_t18 = this.rowexport_type_18[ax].widthloss * this.rowexport_type_18[ax].ttl_marker
      this.sum_end1_2_t18 = this.rowexport_type_18[ax].end1_2 * this.rowexport_type_18[ax].ttl_marker
      this.sum_endless_length_yd_t18 = Number(this.sum_endless_length_yd_t18) + Number(this.rowexport_type_18[ax].endless_length_yd)
      this.sum_spliceloss_t18 = Number(this.sum_spliceloss_t18) + Number(this.rowexport_type_18[ax].spliceloss)
      this.sum_totalcutout_t18 = Number(this.sum_totalcutout_t18) + Number(this.rowexport_type_18[ax].totalcutout)
      this.sum_rement_t18 = Number(this.sum_rement_t18) + Number(this.rowexport_type_18[ax].rement)
      this.summary_plug12_t18 = Number(this.summary_plug12_t18) + Number(this.sum_plug12_t18)
      this.summary_widthloss_t18 = Number(this.summary_widthloss_t18) + Number(this.sum_per_width_loss_t18)
      this.summary_end1_2_t18 = Number(this.summary_end1_2_t18) + Number(this.sum_end1_2_t18)
      this.sum_totallenthloss_t18_first = 0
      this.sum_totallenthloss_t18_first = Number(this.rowexport_type_18[ax].endless_length_yd) + Number(this.rowexport_type_18[ax].spliceloss)
      +Number(this.rowexport_type_18[ax].totalcutout) + Number(this.rowexport_type_18[ax].rement)
      this.sum_totallenthloss_t18 = Number(this.sum_totallenthloss_t18) + Number(this.sum_totallenthloss_t18_first)

      this.sum_totallenthloss_per_t18_first = 0
      this.sum_totallenthloss_per_t18_first = Number(this.rowexport_type_18[ax].avg_end) + Number(this.rowexport_type_18[ax].splicelossper)
      +Number(this.rowexport_type_18[ax].totalcutoutper) + Number(this.rowexport_type_18[ax].rement_per)
      this.sum_totallenthloss_per_t18 = Number(this.sum_totallenthloss_t18) + Number(this.sum_totallenthloss_per_t18_first)
      this.sum_number_1_end_t18 = 0 
      this.sum_number_1_end_t18 = this.rowexport_type_18[ax].endless_length_yd * 36
      this.sum_number_2_end_t18 = 0
      this.sum_number_2_end_t18 = this.rowexport_type_18[ax].ttl_marker/this.rowexport_type_18[ax].length_ydb
      this.sum_plug1_2_end_t18 = 0
      this.sum_plug1_2_end_t18 = this.sum_number_1_end_t18/this.sum_number_2_end_t18

      this.sum_plug1_2_end_total_t18 = 0
     
      this.sum_plug1_2_end_total_t18 = this.sum_plug1_2_end_t18 * this.rowexport_type_18[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t18 = Number(this.sum_all_plug1_2_end_total_t18) + Number(this.sum_plug1_2_end_total_t18)

      this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t18
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t18
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t18
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t18
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t18
      })
       
      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t18,
        value_table:18
      }) 
      

      ws.getCell("N"+count18).value = Number(this.sum_plug1_2_end_t18)
      ws.getCell("N"+count18).numFmt = '0.00';


      count18++
         }
         if(this.sum_ttl_marker_t18 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t18
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t18 = this.summary_plug12_t18 / this.sum_ttl_marker_t18
        if(this.sum_plug12_t18 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t18
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t18 =  this.summary_widthloss_t18 / this.sum_ttl_marker_t18
        if(this.sum_per_width_loss_t18 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t18
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t18 =  this.summary_end1_2_t18/ this.sum_ttl_marker_t18 //M
      if(this.sum_end1_2_t18 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t18
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t18 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t18
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t18 = this.sum_endless_length_yd_t18 / this.sum_ttl_marker_t18 *100 //P
      if(this.sum_avg_end_t18 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t18
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t18 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t18
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t18 = this.sum_spliceloss_t18 / this.sum_ttl_marker_t18 *100
      if(this.sum_splicelossper_t18 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t18
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t18 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t18
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t18 = this.sum_totalcutout_t18 / this.sum_ttl_marker_t18 *100
      if(this.sum_totalcutoutper_t18 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t18
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t18 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t18
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t18 = this.sum_rement_t18 / this.sum_ttl_marker_t18 *100
  if(this.sum_rement_per_t18 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t18
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t18
    if(this.sum_totallenthloss_t18 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t18
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert18 = this.sum_totallenthlosst18 / this.sum_ttl_marker_t18
    if(this.sum_totallenthlosspert18 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert18
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

ws.getCell("A"+countx18).value = "Total Table A18"  
      ws.mergeCells("A"+countx18+":"+"I"+countx18)
      
      if(this.sum_ttl_marker_t18 > 0 && isNaN(this.sum_ttl_marker_t18)==false && isFinite(this.sum_ttl_marker_t18)==true){
      ws.getCell("J"+countx18).value = Number(this.sum_ttl_marker_t18)
      ws.getCell("J"+countx18).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx18).value = 0
      ws.getCell("J"+countx18).numFmt = '0.00';
      }

      if(this.sum_plug12_t18 > 0 && isNaN(this.sum_plug12_t18)==false && isFinite(this.sum_plug12_t18)==true){
      ws.getCell("K"+countx18).value = Number(this.sum_plug12_t18)
      ws.getCell("K"+countx18).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx18).value = 0
      ws.getCell("K"+countx18).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t18 > 0 && isNaN(this.sum_per_width_loss_t18)==false && isFinite(this.sum_per_width_loss_t18)==true){
      ws.getCell("L"+countx18).value = this.sum_per_width_loss_t18/100
      ws.getCell("L"+countx18).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx18).value = 0.00/100
      ws.getCell("L"+countx18).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t18 = 0
      this.total_result_plug1_2_t18 = this.sum_all_plug1_2_end_total_t18 / this.sum_ttl_marker_t18

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t18
      })
      if(this.total_result_plug1_2_t18 > 0 && isNaN(this.total_result_plug1_2_t18)==false && isFinite(this.total_result_plug1_2_t18)==true){
      ws.getCell("N"+countx18).value = Number(this.total_result_plug1_2_t18)
      ws.getCell("N"+countx18).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx18).value = 0.00
      ws.getCell("N"+countx18).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_t18 > 0 && isNaN(this.sum_endless_length_yd_t18)==false && isFinite(this.sum_endless_length_yd_t18)==true){
      ws.getCell("O"+countx18).value = Number(this.sum_endless_length_yd_t18)
      ws.getCell("O"+countx18).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx18).value = 0.00
      ws.getCell("O"+countx18).numFmt = '0.00';
      }

      if(this.sum_avg_end_t18 > 0 && isNaN(this.sum_avg_end_t18)==false && isFinite(this.sum_avg_end_t18)==true){
      ws.getCell("P"+countx18).value = this.sum_avg_end_t18/100
      ws.getCell("P"+countx18).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx18).value = 0/100
      ws.getCell("P"+countx18).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t18 > 0 && isNaN(this.sum_spliceloss_t18)==false && isFinite(this.sum_spliceloss_t18)==true){
      ws.getCell("R"+countx18).value = Number(this.sum_spliceloss_t18)
      ws.getCell("R"+countx18).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx18).value = 0.00
      ws.getCell("R"+countx18).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t18 > 0 && isNaN(this.sum_splicelossper_t18)==false && isFinite(this.sum_splicelossper_t18)==true){
      ws.getCell("S"+countx18).value = this.sum_splicelossper_t18/100
      ws.getCell("S"+countx18).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx18).value = 0/100
      ws.getCell("S"+countx18).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t18 > 0 && isNaN(this.sum_totalcutout_t18)==false && isFinite(this.sum_totalcutout_t18)==true){
      ws.getCell("U"+countx18).value = Number(this.sum_totalcutout_t18)
      ws.getCell("U"+countx18).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx18).value = 0.00
      ws.getCell("U"+countx18).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t18 > 0 && isNaN(this.sum_totalcutoutper_t18)==false && isFinite(this.sum_totalcutoutper_t18)==true){
      ws.getCell("V"+countx18).value = this.sum_totalcutoutper_t18/100
      ws.getCell("V"+countx18).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx18).value = 0.00/100
      ws.getCell("V"+countx18).numFmt = '0.00%';
      }

       if(this.sum_rement_t18 > 0 && isNaN(this.sum_rement_t18)==false && isFinite(this.sum_rement_t18)==true){
      ws.getCell("X"+countx18).value = Number(this.sum_rement_t18)
      ws.getCell("X"+countx18).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx18).value = 0.00
      ws.getCell("X"+countx18).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t18 > 0 && isNaN(this.sum_rement_per_t18)==false && isFinite(this.sum_rement_per_t18)==true){
      ws.getCell("Y"+countx18).value =  this.sum_rement_per_t18/100
      ws.getCell("Y"+countx18).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx18).value =  0.00/100
      ws.getCell("Y"+countx18).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t18 > 0 && isNaN(this.sum_totallenthloss_t18)==false && isFinite(this.sum_totallenthloss_t18)==true){
      ws.getCell("AA"+countx18).value = Number(this.sum_totallenthloss_t18)
      ws.getCell("AA"+countx18).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx18).value = 0
      ws.getCell("AA"+countx18).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t18 = this.sum_totallenthloss_t18 / this.sum_ttl_marker_t18 *100
      if(this.last_sum_totallenthloss_per_t18 > 0 && isNaN(this.last_sum_totallenthloss_per_t18)==false && isFinite(this.last_sum_totallenthloss_per_t18)==true){
      ws.getCell("AB"+countx18).value = this.last_sum_totallenthloss_per_t18/100
      ws.getCell("AB"+countx18).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx18).value = 0/100
      ws.getCell("AB"+countx18).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx18).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //A19

      if(this.rowexport_type_18.length == 0){
       
         var count19 = count18

       }else{
       var count19 = countx18+1
       }
       
      
      
  
        var countx19 = count19 + 1 +Number(this.rowexport_type_19.length)
      
      this.sum_ttl_marker_t19 = 0
      this.sum_plug12_t19 = 0
      this.sum_per_width_loss_t19 = 0
      this.sum_end1_2_t19 = 0
      this.sum_endless_length_yd_t19= 0
      this.sum_avg_end_t19 = 0
      this.sum_spliceloss_t19 = 0
      this.sum_splicelossper_t19 = 0
      this.sum_totalcutout_t19 = 0
      this.sum_totalcutoutper_t19 = 0
      this.sum_rement_t19 = 0
      this.sum_rement_per_t19 = 0
      this.sum_totallenthloss_t19 =0 
      this.sum_totallenthlossper_t19 = 0
      this.summary_widthloss_t19 = 0
      this.summary_plug12_t19 = 0
      this.summary_end1_2_t19 = 0

      
     
      this.sum_all_plug1_2_end_total_t19 = 0

      if(this.rowexport_type_19.length < 1){
      
      count_row_grand++
      
      }else{
          for (var ax = 0; ax<this.rowexport_type_19.length; ax++){ //length +120
      ws.getCell("A"+count19).value = this.rowexport_type_19[ax].mu_date

      ws.getCell("B"+count19).value = this.rowexport_type_19[ax].gm_table

      ws.getCell("C"+count19).value = Number(this.rowexport_type_19[ax].gm_no_short)
      ws.getCell("C"+count19).numFmt = '0.00';

      ws.getCell("D"+count19).value = this.rowexport_type_19[ax].customer

      ws.getCell("E"+count19).value = this.rowexport_type_19[ax].gm_no

      ws.getCell("F"+count19).value = this.rowexport_type_19[ax].fabric_type

      ws.getCell("G"+count19).value = Number(this.rowexport_type_19[ax].obs_width)
      ws.getCell("G"+count19).numFmt = '0.00';

      ws.getCell("H"+count19).value = Number(this.rowexport_type_19[ax].width_inch)
      ws.getCell("H"+count19).numFmt = '0.00';


      ws.getCell("I"+count19).value = Number(this.rowexport_type_19[ax].length_ydb)
      ws.getCell("I"+count19).numFmt = '0.00';

      ws.getCell("J"+count19).value = Number(this.rowexport_type_19[ax].ttl_marker)
      ws.getCell("J"+count19).numFmt = '0.00';

      ws.getCell("K"+count19).value = Number(this.rowexport_type_19[ax].plug12)
      ws.getCell("K"+count19).numFmt = '0.00';

      ws.getCell("L"+count19).value = this.rowexport_type_19[ax].widthloss/100
      ws.getCell("L"+count19).numFmt = '0.00%';

      ws.getCell("M"+count19).value = this.rowexport_type_19[ax].widthloss_code
    
      ws.getCell("O"+count19).value = Number(this.rowexport_type_19[ax].endless_length_yd)
      ws.getCell("O"+count19).numFmt = '0.00';

      ws.getCell("P"+count19).value = this.rowexport_type_19[ax].avg_end/100
      ws.getCell("P"+count19).numFmt = '0.00%';

      ws.getCell("Q"+count19).value = this.rowexport_type_19[ax].avg_end_code
      

      ws.getCell("R"+count19).value = Number(this.rowexport_type_19[ax].spliceloss)
      ws.getCell("R"+count19).numFmt = '0.00';

      ws.getCell("S"+count19).value = this.rowexport_type_19[ax].splicelossper/100
      ws.getCell("S"+count19).numFmt = '0.00%';

      ws.getCell("T"+count19).value = this.rowexport_type_19[ax].per_splice_loss_code
   

      ws.getCell("U"+count19).value = Number(this.rowexport_type_19[ax].totalcutout)
      ws.getCell("U"+count19).numFmt = '0.00';

      ws.getCell("V"+count19).value = this.rowexport_type_19[ax].totalcutoutper/100
      ws.getCell("V"+count19).numFmt = '0.00%';

      ws.getCell("W"+count19).value = this.rowexport_type_19[ax].total_cut_out_per_code

      ws.getCell("X"+count19).value = Number(this.rowexport_type_19[ax].rement)
      ws.getCell("X"+count19).numFmt = '0.00';

      ws.getCell("Y"+count19).value = this.rowexport_type_19[ax].rement_per/100
      ws.getCell("Y"+count19).numFmt = '0.00%';

      ws.getCell("Z"+count19).value = this.rowexport_type_19[ax].remnants_loss_code

      ws.getCell("AA"+count19).value = Number(this.rowexport_type_19[ax].totallenthloss)
      ws.getCell("AA"+count19).numFmt = '0.00';

      ws.getCell("AB"+count19).value = this.rowexport_type_19[ax].totallenthlossper/100
      ws.getCell("AB"+count19).numFmt = '0.00%';
      
      this.sum_plug12_t19 = 0
      this.sum_ttl_marker_t19 = Number(this.sum_ttl_marker_t19) + Number(this.rowexport_type_19[ax].ttl_marker)
      this.sum_plug12_t19 = this.rowexport_type_19[ax].plug12 * this.rowexport_type_19[ax].ttl_marker
      this.sum_per_width_loss_t19 = this.rowexport_type_19[ax].widthloss * this.rowexport_type_19[ax].ttl_marker
      this.sum_end1_2_t19 = this.rowexport_type_19[ax].end1_2 * this.rowexport_type_19[ax].ttl_marker
      this.sum_endless_length_yd_t19 = Number(this.sum_endless_length_yd_t19) + Number(this.rowexport_type_19[ax].endless_length_yd)
      this.sum_spliceloss_t19 = Number(this.sum_spliceloss_t19) + Number(this.rowexport_type_19[ax].spliceloss)
      this.sum_totalcutout_t19 = Number(this.sum_totalcutout_t19) + Number(this.rowexport_type_19[ax].totalcutout)
      this.sum_rement_t19 = Number(this.sum_rement_t19) + Number(this.rowexport_type_19[ax].rement)
      this.summary_plug12_t19 = Number(this.summary_plug12_t19) + Number(this.sum_plug12_t19)
      this.summary_widthloss_t19 = Number(this.summary_widthloss_t19) + Number(this.sum_per_width_loss_t19)
      this.summary_end1_2_t19 = Number(this.summary_end1_2_t19) + Number(this.sum_end1_2_t19)
      this.sum_totallenthloss_t19_first = 0
      this.sum_totallenthloss_t19_first = Number(this.rowexport_type_19[ax].endless_length_yd) + Number(this.rowexport_type_19[ax].spliceloss)
      +Number(this.rowexport_type_19[ax].totalcutout) + Number(this.rowexport_type_19[ax].rement)
      this.sum_totallenthloss_t19 = Number(this.sum_totallenthloss_t19) + Number(this.sum_totallenthloss_t19_first)

      this.sum_totallenthloss_per_t19_first = 0
      this.sum_totallenthloss_per_t19_first = Number(this.rowexport_type_19[ax].avg_end) + Number(this.rowexport_type_19[ax].splicelossper)
      +Number(this.rowexport_type_19[ax].totalcutoutper) + Number(this.rowexport_type_19[ax].rement_per)
      this.sum_totallenthloss_per_t19 = Number(this.sum_totallenthloss_t19) + Number(this.sum_totallenthloss_per_t19_first)
      this.sum_number_1_end_t19 = 0 
      this.sum_number_1_end_t19 = this.rowexport_type_19[ax].endless_length_yd * 36
      this.sum_number_2_end_t19 = 0
      this.sum_number_2_end_t19 = this.rowexport_type_19[ax].ttl_marker/this.rowexport_type_19[ax].length_ydb
      this.sum_plug1_2_end_t19 = 0
      this.sum_plug1_2_end_t19 = this.sum_number_1_end_t19/this.sum_number_2_end_t19

      this.sum_plug1_2_end_total_t19 = 0
     
      this.sum_plug1_2_end_total_t19 = this.sum_plug1_2_end_t19 * this.rowexport_type_19[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t19 = Number(this.sum_all_plug1_2_end_total_t19) + Number(this.sum_plug1_2_end_total_t19)

       this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t19
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t19
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t19
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t19
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t19
      })
       
      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t19,
        value_table:19
      }) 

      ws.getCell("N"+count19).value = Number(this.sum_plug1_2_end_t19)
      ws.getCell("N"+count19).numFmt = '0.00';


      count19++
         }
         if(this.sum_ttl_marker_t19 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t19
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t19 = this.summary_plug12_t19 / this.sum_ttl_marker_t19
        if(this.sum_plug12_t19 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t19
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t19 =  this.summary_widthloss_t19 / this.sum_ttl_marker_t19
        if(this.sum_per_width_loss_t19 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t19
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t19 =  this.summary_end1_2_t19/ this.sum_ttl_marker_t19 //M
      if(this.sum_end1_2_t19 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t19
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t19 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t19
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t19 = this.sum_endless_length_yd_t19 / this.sum_ttl_marker_t19 *100 //P
      if(this.sum_avg_end_t19 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t19
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t19 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t19
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t19 = this.sum_spliceloss_t19 / this.sum_ttl_marker_t19 *100
      if(this.sum_splicelossper_t19 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t19
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t19 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t19
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t19 = this.sum_totalcutout_t19 / this.sum_ttl_marker_t19 *100
      if(this.sum_totalcutoutper_t19 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t19
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t19 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t19
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t19 = this.sum_rement_t19 / this.sum_ttl_marker_t19 *100
  if(this.sum_rement_per_t19 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t19
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t19
    if(this.sum_totallenthloss_t19 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t19
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert19 = this.sum_totallenthlosst19 / this.sum_ttl_marker_t19
    if(this.sum_totallenthlosspert19 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert19
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

ws.getCell("A"+countx19).value = "Total Table A19"  
      ws.mergeCells("A"+countx19+":"+"I"+countx19)
      
      if(this.sum_ttl_marker_t19 > 0 && isNaN(this.sum_ttl_marker_t19)==false && isFinite(this.sum_ttl_marker_t19)==true){
      ws.getCell("J"+countx19).value = Number(this.sum_ttl_marker_t19)
      ws.getCell("J"+countx19).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx19).value = 0
      ws.getCell("J"+countx19).numFmt = '0.00';
      }

      if(this.sum_plug12_t19 > 0 && isNaN(this.sum_plug12_t19)==false && isFinite(this.sum_plug12_t19)==true){
      ws.getCell("K"+countx19).value = Number(this.sum_plug12_t19)
      ws.getCell("K"+countx19).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx19).value = 0
      ws.getCell("K"+countx19).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t19 > 0 && isNaN(this.sum_per_width_loss_t19)==false && isFinite(this.sum_per_width_loss_t19)==true){
      ws.getCell("L"+countx19).value = this.sum_per_width_loss_t19/100
      ws.getCell("L"+countx19).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx19).value = 0.00/100
      ws.getCell("L"+countx19).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t19 = 0
      this.total_result_plug1_2_t19 = this.sum_all_plug1_2_end_total_t19 / this.sum_ttl_marker_t19

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t19
      })
      if(this.total_result_plug1_2_t19 > 0 && isNaN(this.total_result_plug1_2_t19)==false && isFinite(this.total_result_plug1_2_t19)==true){
      ws.getCell("N"+countx19).value = Number(this.total_result_plug1_2_t19)
      ws.getCell("N"+countx19).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx19).value = 0.00
      ws.getCell("N"+countx19).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_t19 > 0 && isNaN(this.sum_endless_length_yd_t19)==false && isFinite(this.sum_endless_length_yd_t19)==true){
      ws.getCell("O"+countx19).value = Number(this.sum_endless_length_yd_t19)
      ws.getCell("O"+countx19).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx19).value = 0.00
      ws.getCell("O"+countx19).numFmt = '0.00';
      }

      if(this.sum_avg_end_t19 > 0 && isNaN(this.sum_avg_end_t19)==false && isFinite(this.sum_avg_end_t19)==true){
      ws.getCell("P"+countx19).value = this.sum_avg_end_t19/100
      ws.getCell("P"+countx19).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx19).value = 0/100
      ws.getCell("P"+countx19).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t19 > 0 && isNaN(this.sum_spliceloss_t19)==false && isFinite(this.sum_spliceloss_t19)==true){
      ws.getCell("R"+countx19).value = Number(this.sum_spliceloss_t19)
      ws.getCell("R"+countx19).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx19).value = 0.00
      ws.getCell("R"+countx19).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t19 > 0 && isNaN(this.sum_splicelossper_t19)==false && isFinite(this.sum_splicelossper_t19)==true){
      ws.getCell("S"+countx19).value = this.sum_splicelossper_t19/100
      ws.getCell("S"+countx19).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx19).value = 0/100
      ws.getCell("S"+countx19).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t19 > 0 && isNaN(this.sum_totalcutout_t19)==false && isFinite(this.sum_totalcutout_t19)==true){
      ws.getCell("U"+countx19).value = Number(this.sum_totalcutout_t19)
      ws.getCell("U"+countx19).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx19).value = 0.00
      ws.getCell("U"+countx19).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t19 > 0 && isNaN(this.sum_totalcutoutper_t19)==false && isFinite(this.sum_totalcutoutper_t19)==true){
      ws.getCell("V"+countx19).value = this.sum_totalcutoutper_t19/100
      ws.getCell("V"+countx19).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx19).value = 0.00/100
      ws.getCell("V"+countx19).numFmt = '0.00%';
      }

       if(this.sum_rement_t19 > 0 && isNaN(this.sum_rement_t19)==false && isFinite(this.sum_rement_t19)==true){
      ws.getCell("X"+countx19).value = Number(this.sum_rement_t19)
      ws.getCell("X"+countx19).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx19).value = 0.00
      ws.getCell("X"+countx19).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t19 > 0 && isNaN(this.sum_rement_per_t19)==false && isFinite(this.sum_rement_per_t19)==true){
      ws.getCell("Y"+countx19).value =  this.sum_rement_per_t19/100
      ws.getCell("Y"+countx19).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx19).value =  0.00/100
      ws.getCell("Y"+countx19).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t19 > 0 && isNaN(this.sum_totallenthloss_t19)==false && isFinite(this.sum_totallenthloss_t19)==true){
      ws.getCell("AA"+countx19).value = Number(this.sum_totallenthloss_t19)
      ws.getCell("AA"+countx19).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx19).value = 0
      ws.getCell("AA"+countx19).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t19 = this.sum_totallenthloss_t19 / this.sum_ttl_marker_t19 *100
      if(this.last_sum_totallenthloss_per_t19 > 0 && isNaN(this.last_sum_totallenthloss_per_t19)==false && isFinite(this.last_sum_totallenthloss_per_t19)==true){
      ws.getCell("AB"+countx19).value = this.last_sum_totallenthloss_per_t19/100
      ws.getCell("AB"+countx19).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx19).value = 0/100
      ws.getCell("AB"+countx19).numFmt = '0.00%'; 
      }
      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx19).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //A20

      if(this.rowexport_type_19.length == 0){
       
         var count20 = count19

       }else{
       var count20 = countx19+1
       }
      
  
        var countx20 = count20 + 1 +Number(this.rowexport_type_20.length)
      
      this.sum_ttl_marker_t20 = 0
      this.sum_plug12_t20 = 0
      this.sum_per_width_loss_t20 = 0
      this.sum_end1_2_t20 = 0
      this.sum_endless_length_yd_t20= 0
      this.sum_avg_end_t20 = 0
      this.sum_spliceloss_t20 = 0
      this.sum_splicelossper_t20 = 0
      this.sum_totalcutout_t20 = 0
      this.sum_totalcutoutper_t20 = 0
      this.sum_rement_t20 = 0
      this.sum_rement_per_t20 = 0
      this.sum_totallenthloss_t20 =0 
      this.sum_totallenthlossper_t20 = 0
      this.summary_widthloss_t20 = 0
      this.summary_plug12_t20 = 0
      this.summary_end1_2_t20 = 0

      
     
      this.sum_all_plug1_2_end_total_t20 = 0
      if(this.rowexport_type_20.length < 1){
      
      count_row_grand++
      
      }else{
          for (var ax = 0; ax<this.rowexport_type_20.length; ax++){ //length +120
      ws.getCell("A"+count20).value = this.rowexport_type_20[ax].mu_date

      ws.getCell("B"+count20).value = this.rowexport_type_20[ax].gm_table

      ws.getCell("C"+count20).value = Number(this.rowexport_type_20[ax].gm_no_short)
      ws.getCell("C"+count20).numFmt = '0.00';

      ws.getCell("D"+count20).value = this.rowexport_type_20[ax].customer

      ws.getCell("E"+count20).value = this.rowexport_type_20[ax].gm_no

      ws.getCell("F"+count20).value = this.rowexport_type_20[ax].fabric_type

      ws.getCell("G"+count20).value = Number(this.rowexport_type_20[ax].obs_width)
      ws.getCell("G"+count20).numFmt = '0.00';

      ws.getCell("H"+count20).value = Number(this.rowexport_type_20[ax].width_inch)
      ws.getCell("H"+count20).numFmt = '0.00';


      ws.getCell("I"+count20).value = Number(this.rowexport_type_20[ax].length_ydb)
      ws.getCell("I"+count20).numFmt = '0.00';

      ws.getCell("J"+count20).value = Number(this.rowexport_type_20[ax].ttl_marker)
      ws.getCell("J"+count20).numFmt = '0.00';

      ws.getCell("K"+count20).value = Number(this.rowexport_type_20[ax].plug12)
      ws.getCell("K"+count20).numFmt = '0.00';

      ws.getCell("L"+count20).value = this.rowexport_type_20[ax].widthloss/100
      ws.getCell("L"+count20).numFmt = '0.00%';

      ws.getCell("M"+count20).value = this.rowexport_type_20[ax].widthloss_code
    
      ws.getCell("O"+count20).value = Number(this.rowexport_type_20[ax].endless_length_yd)
      ws.getCell("O"+count20).numFmt = '0.00';

      ws.getCell("P"+count20).value = this.rowexport_type_20[ax].avg_end/100
      ws.getCell("P"+count20).numFmt = '0.00%';

      ws.getCell("Q"+count20).value = this.rowexport_type_20[ax].avg_end_code
      

      ws.getCell("R"+count20).value = Number(this.rowexport_type_20[ax].spliceloss)
      ws.getCell("R"+count20).numFmt = '0.00';

      ws.getCell("S"+count20).value = this.rowexport_type_20[ax].splicelossper/100
      ws.getCell("S"+count20).numFmt = '0.00%';

      ws.getCell("T"+count20).value = this.rowexport_type_20[ax].per_splice_loss_code
   

      ws.getCell("U"+count20).value = Number(this.rowexport_type_20[ax].totalcutout)
      ws.getCell("U"+count20).numFmt = '0.00';

      ws.getCell("V"+count20).value = this.rowexport_type_20[ax].totalcutoutper/100
      ws.getCell("V"+count20).numFmt = '0.00%';

      ws.getCell("W"+count20).value = this.rowexport_type_20[ax].total_cut_out_per_code

      ws.getCell("X"+count20).value = Number(this.rowexport_type_20[ax].rement)
      ws.getCell("X"+count20).numFmt = '0.00';

      ws.getCell("Y"+count20).value = this.rowexport_type_20[ax].rement_per/100
      ws.getCell("Y"+count20).numFmt = '0.00%';

      ws.getCell("Z"+count20).value = this.rowexport_type_20[ax].remnants_loss_code

      ws.getCell("AA"+count20).value = Number(this.rowexport_type_20[ax].totallenthloss)
      ws.getCell("AA"+count20).numFmt = '0.00';

      ws.getCell("AB"+count20).value = this.rowexport_type_20[ax].totallenthlossper/100
      ws.getCell("AB"+count20).numFmt = '0.00%';
      
      this.sum_plug12_t20 = 0
      this.sum_ttl_marker_t20 = Number(this.sum_ttl_marker_t20) + Number(this.rowexport_type_20[ax].ttl_marker)
      this.sum_plug12_t20 = this.rowexport_type_20[ax].plug12 * this.rowexport_type_20[ax].ttl_marker
      this.sum_per_width_loss_t20 = this.rowexport_type_20[ax].widthloss * this.rowexport_type_20[ax].ttl_marker
      this.sum_end1_2_t20 = this.rowexport_type_20[ax].end1_2 * this.rowexport_type_20[ax].ttl_marker
      this.sum_endless_length_yd_t20 = Number(this.sum_endless_length_yd_t20) + Number(this.rowexport_type_20[ax].endless_length_yd)
      this.sum_spliceloss_t20 = Number(this.sum_spliceloss_t20) + Number(this.rowexport_type_20[ax].spliceloss)
      this.sum_totalcutout_t20 = Number(this.sum_totalcutout_t20) + Number(this.rowexport_type_20[ax].totalcutout)
      this.sum_rement_t20 = Number(this.sum_rement_t20) + Number(this.rowexport_type_20[ax].rement)
      this.summary_plug12_t20 = Number(this.summary_plug12_t20) + Number(this.sum_plug12_t20)
      this.summary_widthloss_t20 = Number(this.summary_widthloss_t20) + Number(this.sum_per_width_loss_t20)
      this.summary_end1_2_t20 = Number(this.summary_end1_2_t20) + Number(this.sum_end1_2_t20)
      this.sum_totallenthloss_t20_first = 0
      this.sum_totallenthloss_t20_first = Number(this.rowexport_type_20[ax].endless_length_yd) + Number(this.rowexport_type_20[ax].spliceloss)
      +Number(this.rowexport_type_20[ax].totalcutout) + Number(this.rowexport_type_20[ax].rement)
      this.sum_totallenthloss_t20 = Number(this.sum_totallenthloss_t20) + Number(this.sum_totallenthloss_t20_first)

      this.sum_totallenthloss_per_t20_first = 0
      this.sum_totallenthloss_per_t20_first = Number(this.rowexport_type_20[ax].avg_end) + Number(this.rowexport_type_20[ax].splicelossper)
      +Number(this.rowexport_type_20[ax].totalcutoutper) + Number(this.rowexport_type_20[ax].rement_per)
      this.sum_totallenthloss_per_t20 = Number(this.sum_totallenthloss_t20) + Number(this.sum_totallenthloss_per_t20_first)
      this.sum_number_1_end_t20 = 0 
      this.sum_number_1_end_t20 = this.rowexport_type_20[ax].endless_length_yd * 36
      this.sum_number_2_end_t20 = 0
      this.sum_number_2_end_t20 = this.rowexport_type_20[ax].ttl_marker/this.rowexport_type_20[ax].length_ydb
      this.sum_plug1_2_end_t20 = 0
      this.sum_plug1_2_end_t20 = this.sum_number_1_end_t20/this.sum_number_2_end_t20

      this.sum_plug1_2_end_total_t20 = 0
     
      this.sum_plug1_2_end_total_t20 = this.sum_plug1_2_end_t20 * this.rowexport_type_20[ax].ttl_marker
      this.sum_all_plug1_2_end_total_t20 = Number(this.sum_all_plug1_2_end_total_t20) + Number(this.sum_plug1_2_end_total_t20)

       this.grand_total_end_loss.push({
        value:this.sum_endless_length_yd_t20
      })
      this.grand_total_splice_loss.push({
        value:this.sum_spliceloss_t20
      })
      this.grand_total_cut_out_loss.push({
        value:this.sum_totalcutout_t20
      })
      this.grand_total_remnants_loss.push({
        value:this.sum_rement_t20
      })
      this.grand_total_length_loss.push({
        value:this.sum_totallenthloss_t20
      })
       
      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_t20,
        value_table:20
      }) 

      ws.getCell("N"+count20).value = Number(this.sum_plug1_2_end_t20)
      ws.getCell("N"+count20).numFmt = '0.00';


      count20++
         }
         if(this.sum_ttl_marker_t20 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_t20
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_t20 = this.summary_plug12_t20 / this.sum_ttl_marker_t20
        if(this.sum_plug12_t20 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_t20
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_t20 =  this.summary_widthloss_t20 / this.sum_ttl_marker_t20
        if(this.sum_per_width_loss_t20 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_t20
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_t20 =  this.summary_end1_2_t20/ this.sum_ttl_marker_t20 //M
      if(this.sum_end1_2_t20 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_t20
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_t20 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_t20
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_t20 = this.sum_endless_length_yd_t20 / this.sum_ttl_marker_t20 *100 //P
      if(this.sum_avg_end_t20 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_t20
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_t20 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_t20
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_t20 = this.sum_spliceloss_t20 / this.sum_ttl_marker_t20 *100
      if(this.sum_splicelossper_t20 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_t20
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_t20 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_t20
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_t20 = this.sum_totalcutout_t20 / this.sum_ttl_marker_t20 *100
      if(this.sum_totalcutoutper_t20 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_t20
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_t20 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_t20
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_t20 = this.sum_rement_t20 / this.sum_ttl_marker_t20 *100
  if(this.sum_rement_per_t20 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_t20
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_t20
    if(this.sum_totallenthloss_t20 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_t20
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspert20 = this.sum_totallenthlosst20 / this.sum_ttl_marker_t20
    if(this.sum_totallenthlosspert20 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspert20
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

ws.getCell("A"+countx20).value = "Total Table A20"  
      ws.mergeCells("A"+countx20+":"+"I"+countx20)
      
      if(this.sum_ttl_marker_t20 > 0 && isNaN(this.sum_ttl_marker_t20)==false && isFinite(this.sum_ttl_marker_t20)==true){
      ws.getCell("J"+countx20).value = Number(this.sum_ttl_marker_t20)
      ws.getCell("J"+countx20).numFmt = '0.00';
      }else{
      ws.getCell("J"+countx20).value = 0
      ws.getCell("J"+countx20).numFmt = '0.00';
      }

      if(this.sum_plug12_t20 > 0 && isNaN(this.sum_plug12_t20)==false && isFinite(this.sum_plug12_t20)==true){
      ws.getCell("K"+countx20).value = Number(this.sum_plug12_t20)
      ws.getCell("K"+countx20).numFmt = '0.00';
      }else{
      ws.getCell("K"+countx20).value = 0
      ws.getCell("K"+countx20).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_t20 > 0 && isNaN(this.sum_per_width_loss_t20)==false && isFinite(this.sum_per_width_loss_t20)==true){
      ws.getCell("L"+countx20).value = this.sum_per_width_loss_t20/100
      ws.getCell("L"+countx20).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countx20).value = 0.00/100
      ws.getCell("L"+countx20).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_t20 = 0
      this.total_result_plug1_2_t20 = this.sum_all_plug1_2_end_total_t20 / this.sum_ttl_marker_t20

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_t20
      })
      if(this.total_result_plug1_2_t20 > 0 && isNaN(this.total_result_plug1_2_t20)==false && isFinite(this.total_result_plug1_2_t20)==true){
      ws.getCell("N"+countx20).value = Number(this.total_result_plug1_2_t20)
      ws.getCell("N"+countx20).numFmt = '0.00';
      }else{
      ws.getCell("N"+countx20).value = 0.00
      ws.getCell("N"+countx20).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_t20 > 0 && isNaN(this.sum_endless_length_yd_t20)==false && isFinite(this.sum_endless_length_yd_t20)==true){
      ws.getCell("O"+countx20).value = Number(this.sum_endless_length_yd_t20)
      ws.getCell("O"+countx20).numFmt = '0.00';
      }else{
      ws.getCell("O"+countx20).value = 0.00
      ws.getCell("O"+countx20).numFmt = '0.00';
      }

      if(this.sum_avg_end_t20 > 0 && isNaN(this.sum_avg_end_t20)==false && isFinite(this.sum_avg_end_t20)==true){
      ws.getCell("P"+countx20).value = this.sum_avg_end_t20/100
      ws.getCell("P"+countx20).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countx20).value = 0/100
      ws.getCell("P"+countx20).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_t20 > 0 && isNaN(this.sum_spliceloss_t20)==false && isFinite(this.sum_spliceloss_t20)==true){
      ws.getCell("R"+countx20).value = Number(this.sum_spliceloss_t20)
      ws.getCell("R"+countx20).numFmt = '0.00';
      }else{
       ws.getCell("R"+countx20).value = 0.00
      ws.getCell("R"+countx20).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_t20 > 0 && isNaN(this.sum_splicelossper_t20)==false && isFinite(this.sum_splicelossper_t20)==true){
      ws.getCell("S"+countx20).value = this.sum_splicelossper_t20/100
      ws.getCell("S"+countx20).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countx20).value = 0/100
      ws.getCell("S"+countx20).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_t20 > 0 && isNaN(this.sum_totalcutout_t20)==false && isFinite(this.sum_totalcutout_t20)==true){
      ws.getCell("U"+countx20).value = Number(this.sum_totalcutout_t20)
      ws.getCell("U"+countx20).numFmt = '0.00';
      }else{
       ws.getCell("U"+countx20).value = 0.00
      ws.getCell("U"+countx20).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_t20 > 0 && isNaN(this.sum_totalcutoutper_t20)==false && isFinite(this.sum_totalcutoutper_t20)==true){
      ws.getCell("V"+countx20).value = this.sum_totalcutoutper_t20/100
      ws.getCell("V"+countx20).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countx20).value = 0.00/100
      ws.getCell("V"+countx20).numFmt = '0.00%';
      }

       if(this.sum_rement_t20 > 0 && isNaN(this.sum_rement_t20)==false && isFinite(this.sum_rement_t20)==true){
      ws.getCell("X"+countx20).value = Number(this.sum_rement_t20)
      ws.getCell("X"+countx20).numFmt = '0.00';
       }else{
      ws.getCell("X"+countx20).value = 0.00
      ws.getCell("X"+countx20).numFmt = '0.00';  
       }

      if(this.sum_rement_per_t20 > 0 && isNaN(this.sum_rement_per_t20)==false && isFinite(this.sum_rement_per_t20)==true){
      ws.getCell("Y"+countx20).value =  this.sum_rement_per_t20/100
      ws.getCell("Y"+countx20).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countx20).value =  0.00/100
      ws.getCell("Y"+countx20).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_t20 > 0 && isNaN(this.sum_totallenthloss_t20)==false && isFinite(this.sum_totallenthloss_t20)==true){
      ws.getCell("AA"+countx20).value = Number(this.sum_totallenthloss_t20)
      ws.getCell("AA"+countx20).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countx20).value = 0
      ws.getCell("AA"+countx20).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_t20 = this.sum_totallenthloss_t20 / this.sum_ttl_marker_t20 *100
      if(this.last_sum_totallenthloss_per_t20 > 0 && isNaN(this.last_sum_totallenthloss_per_t20)==false && isFinite(this.last_sum_totallenthloss_per_t20)==true){
      ws.getCell("AB"+countx20).value = this.last_sum_totallenthloss_per_t20/100
      ws.getCell("AB"+countx20).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countx20).value = 0/100
      ws.getCell("AB"+countx20).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countx20).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }


      //M01

      if(this.rowexport_type_20.length == 0){
       
         var countM1 = count20

       }else{
        var countM1 = countx20+1
       }
     

        var countxM1 = countM1 + 1 +Number(this.rowexport_type_M1.length)
      
      this.sum_ttl_marker_tM1 = 0
      this.sum_plug12_tM1 = 0
      this.sum_per_width_loss_tM1 = 0
      this.sum_end1_2_tM1 = 0
      this.sum_endless_length_yd_tM1= 0
      this.sum_avg_end_tM1 = 0
      this.sum_spliceloss_tM1 = 0
      this.sum_splicelossper_tM1 = 0
      this.sum_totalcutout_tM1 = 0
      this.sum_totalcutoutper_tM1 = 0
      this.sum_rement_tM1 = 0
      this.sum_rement_per_tM1 = 0
      this.sum_totallenthloss_tM1 =0 
      this.sum_totallenthlossper_tM1 = 0
      this.summary_widthloss_tM1 = 0
      this.summary_plug12_tM1 = 0
      this.summary_end1_2_tM1 = 0

      
     this.sum_all_plug1_2_end_total_tM1 = 0
      if(this.rowexport_type_M1.length < 1){
      
      count_row_grand++
      
      }else{
          for (var ax = 0; ax<this.rowexport_type_M1.length; ax++){ //length +M10
      ws.getCell("A"+countM1).value = this.rowexport_type_M1[ax].mu_date

      ws.getCell("B"+countM1).value = this.rowexport_type_M1[ax].gm_table

      ws.getCell("C"+countM1).value = Number(this.rowexport_type_M1[ax].gm_no_short)
      ws.getCell("C"+countM1).numFmt = '0.00';

      ws.getCell("D"+countM1).value = this.rowexport_type_M1[ax].customer

      ws.getCell("E"+countM1).value = this.rowexport_type_M1[ax].gm_no

      ws.getCell("F"+countM1).value = this.rowexport_type_M1[ax].fabric_type

      ws.getCell("G"+countM1).value = Number(this.rowexport_type_M1[ax].obs_width)
      ws.getCell("G"+countM1).numFmt = '0.00';

      ws.getCell("H"+countM1).value = Number(this.rowexport_type_M1[ax].width_inch)
      ws.getCell("H"+countM1).numFmt = '0.00';


      ws.getCell("I"+countM1).value = Number(this.rowexport_type_M1[ax].length_ydb)
      ws.getCell("I"+countM1).numFmt = '0.00';

      ws.getCell("J"+countM1).value = Number(this.rowexport_type_M1[ax].ttl_marker)
      ws.getCell("J"+countM1).numFmt = '0.00';

      ws.getCell("K"+countM1).value = Number(this.rowexport_type_M1[ax].plug12)
      ws.getCell("K"+countM1).numFmt = '0.00';

      ws.getCell("L"+countM1).value = this.rowexport_type_M1[ax].widthloss/100
      ws.getCell("L"+countM1).numFmt = '0.00%';

      ws.getCell("M"+countM1).value = this.rowexport_type_M1[ax].widthloss_code

      ws.getCell("O"+countM1).value = Number(this.rowexport_type_M1[ax].endless_length_yd)
      ws.getCell("O"+countM1).numFmt = '0.00';

      ws.getCell("P"+countM1).value = this.rowexport_type_M1[ax].avg_end/100
      ws.getCell("P"+countM1).numFmt = '0.00%';

      ws.getCell("Q"+countM1).value = this.rowexport_type_M1[ax].avg_end_code
      

      ws.getCell("R"+countM1).value = Number(this.rowexport_type_M1[ax].spliceloss)
      ws.getCell("R"+countM1).numFmt = '0.00';

      ws.getCell("S"+countM1).value = this.rowexport_type_M1[ax].splicelossper/100
      ws.getCell("S"+countM1).numFmt = '0.00%';

      ws.getCell("T"+countM1).value = this.rowexport_type_M1[ax].per_splice_loss_code
   

      ws.getCell("U"+countM1).value = Number(this.rowexport_type_M1[ax].totalcutout)
      ws.getCell("U"+countM1).numFmt = '0.00';

    
      ws.getCell("V"+countM1).value = this.rowexport_type_M1[ax].totalcutoutper/100
      ws.getCell("V"+countM1).numFmt = '0.00%';

      ws.getCell("W"+countM1).value = this.rowexport_type_M1[ax].total_cut_out_per_code

      ws.getCell("X"+countM1).value = Number(this.rowexport_type_M1[ax].rement)
      ws.getCell("X"+countM1).numFmt = '0.00';

      ws.getCell("Y"+countM1).value = this.rowexport_type_M1[ax].rement_per/100
      ws.getCell("Y"+countM1).numFmt = '0.00%';

      ws.getCell("Z"+countM1).value = this.rowexport_type_M1[ax].remnants_loss_code

      ws.getCell("AA"+countM1).value = Number(this.rowexport_type_M1[ax].totallenthloss)
      ws.getCell("AA"+countM1).numFmt = '0.00';

      ws.getCell("AB"+countM1).value = this.rowexport_type_M1[ax].totallenthlossper/100
      ws.getCell("AB"+countM1).numFmt = '0.00%';
      
      this.sum_plug12_tM1 = 0
      this.sum_ttl_marker_tM1 = Number(this.sum_ttl_marker_tM1) + Number(this.rowexport_type_M1[ax].ttl_marker)
      this.sum_plug12_tM1 = this.rowexport_type_M1[ax].plug12 * this.rowexport_type_M1[ax].ttl_marker
      this.sum_per_width_loss_tM1 = this.rowexport_type_M1[ax].widthloss * this.rowexport_type_M1[ax].ttl_marker
      this.sum_end1_2_tM1 = this.rowexport_type_M1[ax].end1_2 * this.rowexport_type_M1[ax].ttl_marker
      this.sum_endless_length_yd_tM1 = Number(this.sum_endless_length_yd_tM1) + Number(this.rowexport_type_M1[ax].endless_length_yd)
      this.sum_spliceloss_tM1 = Number(this.sum_spliceloss_tM1) + Number(this.rowexport_type_M1[ax].spliceloss)
      this.sum_totalcutout_tM1 = Number(this.sum_totalcutout_tM1) + Number(this.rowexport_type_M1[ax].totalcutout)
      this.sum_rement_tM1 = Number(this.sum_rement_tM1) + Number(this.rowexport_type_M1[ax].rement)
      this.summary_plug12_tM1 = Number(this.summary_plug12_tM1) + Number(this.sum_plug12_tM1)
      this.summary_widthloss_tM1 = Number(this.summary_widthloss_tM1) + Number(this.sum_per_width_loss_tM1)
      this.summary_end1_2_tM1 = Number(this.summary_end1_2_tM1) + Number(this.sum_end1_2_tM1)
      this.sum_totallenthloss_tM1_first = 0
      this.sum_totallenthloss_tM1_first = Number(this.rowexport_type_M1[ax].endless_length_yd) + Number(this.rowexport_type_M1[ax].spliceloss)
      +Number(this.rowexport_type_M1[ax].totalcutout) + Number(this.rowexport_type_M1[ax].rement)
      this.sum_totallenthloss_tM1 = Number(this.sum_totallenthloss_tM1) + Number(this.sum_totallenthloss_tM1_first)

      this.sum_totallenthloss_per_tM1_first = 0
      this.sum_totallenthloss_per_tM1_first = Number(this.rowexport_type_M1[ax].avg_end) + Number(this.rowexport_type_M1[ax].splicelossper)
      +Number(this.rowexport_type_M1[ax].totalcutoutper) + Number(this.rowexport_type_M1[ax].rement_per)
      this.sum_totallenthloss_per_tM1 = Number(this.sum_totallenthloss_tM1) + Number(this.sum_totallenthloss_per_tM1_first)
      this.sum_number_1_end_tM1 = 0 
      this.sum_number_1_end_tM1 = this.rowexport_type_M1[ax].endless_length_yd * 36
      this.sum_number_2_end_tM1 = 0
      this.sum_number_2_end_tM1 = this.rowexport_type_M1[ax].ttl_marker/this.rowexport_type_M1[ax].length_ydb
      this.sum_plug1_2_end_tM1 = 0
      this.sum_plug1_2_end_tM1 = this.sum_number_1_end_tM1/this.sum_number_2_end_tM1

      this.sum_plug1_2_end_total_tM1 = 0
     
      this.sum_plug1_2_end_total_tM1 = this.sum_plug1_2_end_tM1 * this.rowexport_type_M1[ax].ttl_marker
      this.sum_all_plug1_2_end_total_tM1 = Number(this.sum_all_plug1_2_end_total_tM1) + Number(this.sum_plug1_2_end_total_tM1)

      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_tM1,
        value_table:"M1"
      })

       
      ws.getCell("N"+countM1).value = Number(this.sum_plug1_2_end_tM1)
      ws.getCell("N"+countM1).numFmt = '0.00';

      countM1++
         }
         if(this.sum_ttl_marker_tM1 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_tM1
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_tM1 = this.summary_plug12_tM1 / this.sum_ttl_marker_tM1
        if(this.sum_plug12_tM1 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_tM1
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_tM1 =  this.summary_widthloss_tM1 / this.sum_ttl_marker_tM1
        if(this.sum_per_width_loss_tM1 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_tM1
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_tM1 =  this.summary_end1_2_tM1/ this.sum_ttl_marker_tM1 //M
      if(this.sum_end1_2_tM1 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_tM1
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_tM1 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_tM1
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_tM1 = this.sum_endless_length_yd_tM1 / this.sum_ttl_marker_tM1 *100 //P
      if(this.sum_avg_end_tM1 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_tM1
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_tM1 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_tM1
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_tM1 = this.sum_spliceloss_tM1 / this.sum_ttl_marker_tM1 *100
      if(this.sum_splicelossper_tM1 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_tM1
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_tM1 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_tM1
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_tM1 = this.sum_totalcutout_tM1 / this.sum_ttl_marker_tM1 *100
      if(this.sum_totalcutoutper_tM1 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_tM1
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_tM1 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_tM1
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_tM1 = this.sum_rement_tM1 / this.sum_ttl_marker_tM1 *100
  if(this.sum_rement_per_tM1 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_tM1
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_tM1
    if(this.sum_totallenthloss_tM1 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_tM1
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspertM1 = this.sum_totallenthlosstM1 / this.sum_ttl_marker_tM1
    if(this.sum_totallenthlosspertM1 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspertM1
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }


     ws.getCell("A"+countxM1).value = "Total Table M1"  
      ws.mergeCells("A"+countxM1+":"+"I"+countxM1)
      
      if(this.sum_ttl_marker_tM1 > 0 && isNaN(this.sum_ttl_marker_tM1)==false && isFinite(this.sum_ttl_marker_tM1)==true){
      ws.getCell("J"+countxM1).value = Number(this.sum_ttl_marker_tM1)
      ws.getCell("J"+countxM1).numFmt = '0.00';
      }else{
      ws.getCell("J"+countxM1).value = 0
      ws.getCell("J"+countxM1).numFmt = '0.00';
      }

      if(this.sum_plug12_tM1 > 0 && isNaN(this.sum_plug12_tM1)==false && isFinite(this.sum_plug12_tM1)==true){
      ws.getCell("K"+countxM1).value = Number(this.sum_plug12_tM1)
      ws.getCell("K"+countxM1).numFmt = '0.00';
      }else{
      ws.getCell("K"+countxM1).value = 0
      ws.getCell("K"+countxM1).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_tM1 > 0 && isNaN(this.sum_per_width_loss_tM1)==false && isFinite(this.sum_per_width_loss_tM1)==true){
      ws.getCell("L"+countxM1).value = this.sum_per_width_loss_tM1/100
      ws.getCell("L"+countxM1).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countxM1).value = 0.00/100
      ws.getCell("L"+countxM1).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_tM1 = 0
      this.total_result_plug1_2_tM1 = this.sum_all_plug1_2_end_total_tM1 / this.sum_ttl_marker_tM1

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_tM1
      })
      if(this.total_result_plug1_2_tM1 > 0 && isNaN(this.total_result_plug1_2_tM1)==false && isFinite(this.total_result_plug1_2_tM1)==true){
      ws.getCell("N"+countxM1).value = Number(this.total_result_plug1_2_tM1)
      ws.getCell("N"+countxM1).numFmt = '0.00';
      }else{
      ws.getCell("N"+countxM1).value = 0.00
      ws.getCell("N"+countxM1).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_tM1 > 0 && isNaN(this.sum_endless_length_yd_tM1)==false && isFinite(this.sum_endless_length_yd_tM1)==true){
      ws.getCell("O"+countxM1).value = Number(this.sum_endless_length_yd_tM1)
      ws.getCell("O"+countxM1).numFmt = '0.00';
      }else{
      ws.getCell("O"+countxM1).value = 0.00
      ws.getCell("O"+countxM1).numFmt = '0.00';
      }

      if(this.sum_avg_end_tM1 > 0 && isNaN(this.sum_avg_end_tM1)==false && isFinite(this.sum_avg_end_tM1)==true){
      ws.getCell("P"+countxM1).value = this.sum_avg_end_tM1/100
      ws.getCell("P"+countxM1).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countxM1).value = 0/100
      ws.getCell("P"+countxM1).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_tM1 > 0 && isNaN(this.sum_spliceloss_tM1)==false && isFinite(this.sum_spliceloss_tM1)==true){
      ws.getCell("R"+countxM1).value = Number(this.sum_spliceloss_tM1)
      ws.getCell("R"+countxM1).numFmt = '0.00';
      }else{
       ws.getCell("R"+countxM1).value = 0.00
      ws.getCell("R"+countxM1).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_tM1 > 0 && isNaN(this.sum_splicelossper_tM1)==false && isFinite(this.sum_splicelossper_tM1)==true){
      ws.getCell("S"+countxM1).value = this.sum_splicelossper_tM1/100
      ws.getCell("S"+countxM1).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countxM1).value = 0/100
      ws.getCell("S"+countxM1).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_tM1 > 0 && isNaN(this.sum_totalcutout_tM1)==false && isFinite(this.sum_totalcutout_tM1)==true){
      ws.getCell("U"+countxM1).value = Number(this.sum_totalcutout_tM1)
      ws.getCell("U"+countxM1).numFmt = '0.00';
      }else{
       ws.getCell("U"+countxM1).value = 0.00
      ws.getCell("U"+countxM1).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_tM1 > 0 && isNaN(this.sum_totalcutoutper_tM1)==false && isFinite(this.sum_totalcutoutper_tM1)==true){
      ws.getCell("V"+countxM1).value = this.sum_totalcutoutper_tM1/100
      ws.getCell("V"+countxM1).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countxM1).value = 0.00/100
      ws.getCell("V"+countxM1).numFmt = '0.00%';
      }

       if(this.sum_rement_tM1 > 0 && isNaN(this.sum_rement_tM1)==false && isFinite(this.sum_rement_tM1)==true){
      ws.getCell("X"+countxM1).value = Number(this.sum_rement_tM1)
      ws.getCell("X"+countxM1).numFmt = '0.00';
       }else{
      ws.getCell("X"+countxM1).value = 0.00
      ws.getCell("X"+countxM1).numFmt = '0.00';  
       }

      if(this.sum_rement_per_tM1 > 0 && isNaN(this.sum_rement_per_tM1)==false && isFinite(this.sum_rement_per_tM1)==true){
      ws.getCell("Y"+countxM1).value =  this.sum_rement_per_tM1/100
      ws.getCell("Y"+countxM1).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countxM1).value =  0.00/100
      ws.getCell("Y"+countxM1).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_tM1 > 0 && isNaN(this.sum_totallenthloss_tM1)==false && isFinite(this.sum_totallenthloss_tM1)==true){
      ws.getCell("AA"+countxM1).value = Number(this.sum_totallenthloss_tM1)
      ws.getCell("AA"+countxM1).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countxM1).value = 0
      ws.getCell("AA"+countxM1).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_tM1 = this.sum_totallenthloss_tM1 / this.sum_ttl_marker_tM1 *100
      if(this.last_sum_totallenthloss_per_tM1 > 0 && isNaN(this.last_sum_totallenthloss_per_tM1)==false && isFinite(this.last_sum_totallenthloss_per_tM1)==true){
      ws.getCell("AB"+countxM1).value = this.last_sum_totallenthloss_per_tM1/100
      ws.getCell("AB"+countxM1).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countxM1).value = 0/100
      ws.getCell("AB"+countxM1).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countxM1).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //M2

      if(this.rowexport_type_M1.length == 0){
       
         var countM2 = countM1

       }else{
        var countM2 = countxM1+1
       }
       

        var countxM2 = countM2 + 1 +Number(this.rowexport_type_M2.length)
      
      this.sum_ttl_marker_tM2 = 0
      this.sum_plug12_tM2 = 0
      this.sum_per_width_loss_tM2 = 0
      this.sum_end1_2_tM2 = 0
      this.sum_endless_length_yd_tM2= 0
      this.sum_avg_end_tM2 = 0
      this.sum_spliceloss_tM2 = 0
      this.sum_splicelossper_tM2 = 0
      this.sum_totalcutout_tM2 = 0
      this.sum_totalcutoutper_tM2 = 0
      this.sum_rement_tM2 = 0
      this.sum_rement_per_tM2 = 0
      this.sum_totallenthloss_tM2 =0 
      this.sum_totallenthlossper_tM2 = 0
      this.summary_widthloss_tM2 = 0
      this.summary_plug12_tM2 = 0
      this.summary_end1_2_tM2 = 0

      this.sum_all_plug1_2_end_total_tM2 = 0
     
    if(this.rowexport_type_M2.length < 1){
      
      count_row_grand++
      
      }else{
          for (var ax = 0; ax<this.rowexport_type_M2.length; ax++){ //length +M20
      ws.getCell("A"+countM2).value = this.rowexport_type_M2[ax].mu_date

      ws.getCell("B"+countM2).value = this.rowexport_type_M2[ax].gm_table

      ws.getCell("C"+countM2).value = Number(this.rowexport_type_M2[ax].gm_no_short)
      ws.getCell("C"+countM2).numFmt = '0.00';

      ws.getCell("D"+countM2).value = this.rowexport_type_M2[ax].customer

      ws.getCell("E"+countM2).value = this.rowexport_type_M2[ax].gm_no

      ws.getCell("F"+countM2).value = this.rowexport_type_M2[ax].fabric_type

      ws.getCell("G"+countM2).value = Number(this.rowexport_type_M2[ax].obs_width)
      ws.getCell("G"+countM2).numFmt = '0.00';

      ws.getCell("H"+countM2).value = Number(this.rowexport_type_M2[ax].width_inch)
      ws.getCell("H"+countM2).numFmt = '0.00';


      ws.getCell("I"+countM2).value = Number(this.rowexport_type_M2[ax].length_ydb)
      ws.getCell("I"+countM2).numFmt = '0.00';

      ws.getCell("J"+countM2).value = Number(this.rowexport_type_M2[ax].ttl_marker)
      ws.getCell("J"+countM2).numFmt = '0.00';

      ws.getCell("K"+countM2).value = Number(this.rowexport_type_M2[ax].plug12)
      ws.getCell("K"+countM2).numFmt = '0.00';

      ws.getCell("L"+countM2).value = this.rowexport_type_M2[ax].widthloss/100
      ws.getCell("L"+countM2).numFmt = '0.00%';

      ws.getCell("M"+countM2).value = this.rowexport_type_M2[ax].widthloss_code
    

      ws.getCell("O"+countM2).value = Number(this.rowexport_type_M2[ax].endless_length_yd)
      ws.getCell("O"+countM2).numFmt = '0.00';

      ws.getCell("P"+countM2).value = this.rowexport_type_M2[ax].avg_end/100
      ws.getCell("P"+countM2).numFmt = '0.00%';

      ws.getCell("Q"+countM2).value = this.rowexport_type_M2[ax].avg_end_code
      

      ws.getCell("R"+countM2).value = Number(this.rowexport_type_M2[ax].spliceloss)
      ws.getCell("R"+countM2).numFmt = '0.00';

      ws.getCell("S"+countM2).value = this.rowexport_type_M2[ax].splicelossper/100
      ws.getCell("S"+countM2).numFmt = '0.00%';

      ws.getCell("T"+countM2).value = this.rowexport_type_M2[ax].per_splice_loss_code
   

      ws.getCell("U"+countM2).value = Number(this.rowexport_type_M2[ax].totalcutout)
      ws.getCell("U"+countM2).numFmt = '0.00';

      ws.getCell("V"+countM2).value = this.rowexport_type_M2[ax].totalcutoutper/100
      ws.getCell("V"+countM2).numFmt = '0.00%';

      ws.getCell("W"+countM2).value = this.rowexport_type_M2[ax].total_cut_out_per_code

      ws.getCell("X"+countM2).value = Number(this.rowexport_type_M2[ax].rement)
      ws.getCell("X"+countM2).numFmt = '0.00';

      ws.getCell("Y"+countM2).value = this.rowexport_type_M2[ax].rement_per/100
      ws.getCell("Y"+countM2).numFmt = '0.00%';

      ws.getCell("Z"+countM2).value = this.rowexport_type_M2[ax].remnants_loss_code

      ws.getCell("AA"+countM2).value = Number(this.rowexport_type_M2[ax].totallenthloss)
      ws.getCell("AA"+countM2).numFmt = '0.00';

      ws.getCell("AB"+countM2).value = this.rowexport_type_M2[ax].totallenthlossper/100
      ws.getCell("AB"+countM2).numFmt = '0.00%';
      
      this.sum_plug12_tM2 = 0
      this.sum_ttl_marker_tM2 = Number(this.sum_ttl_marker_tM2) + Number(this.rowexport_type_M2[ax].ttl_marker)
      this.sum_plug12_tM2 = this.rowexport_type_M2[ax].plug12 * this.rowexport_type_M2[ax].ttl_marker
      this.sum_per_width_loss_tM2 = this.rowexport_type_M2[ax].widthloss * this.rowexport_type_M2[ax].ttl_marker
      this.sum_end1_2_tM2 = this.rowexport_type_M2[ax].end1_2 * this.rowexport_type_M2[ax].ttl_marker
      this.sum_endless_length_yd_tM2 = Number(this.sum_endless_length_yd_tM2) + Number(this.rowexport_type_M2[ax].endless_length_yd)
      this.sum_spliceloss_tM2 = Number(this.sum_spliceloss_tM2) + Number(this.rowexport_type_M2[ax].spliceloss)
      this.sum_totalcutout_tM2 = Number(this.sum_totalcutout_tM2) + Number(this.rowexport_type_M2[ax].totalcutout)
      this.sum_rement_tM2 = Number(this.sum_rement_tM2) + Number(this.rowexport_type_M2[ax].rement)
      this.summary_plug12_tM2 = Number(this.summary_plug12_tM2) + Number(this.sum_plug12_tM2)
      this.summary_widthloss_tM2 = Number(this.summary_widthloss_tM2) + Number(this.sum_per_width_loss_tM2)
      this.summary_end1_2_tM2 = Number(this.summary_end1_2_tM2) + Number(this.sum_end1_2_tM2)
      this.sum_totallenthloss_tM2_first = 0
      this.sum_totallenthloss_tM2_first = Number(this.rowexport_type_M2[ax].endless_length_yd) + Number(this.rowexport_type_M2[ax].spliceloss)
      +Number(this.rowexport_type_M2[ax].totalcutout) + Number(this.rowexport_type_M2[ax].rement)
      this.sum_totallenthloss_tM2 = Number(this.sum_totallenthloss_tM2) + Number(this.sum_totallenthloss_tM2_first)

      this.sum_totallenthloss_per_tM2_first = 0
      this.sum_totallenthloss_per_tM2_first = Number(this.rowexport_type_M2[ax].avg_end) + Number(this.rowexport_type_M2[ax].splicelossper)
      +Number(this.rowexport_type_M2[ax].totalcutoutper) + Number(this.rowexport_type_M2[ax].rement_per)
      this.sum_totallenthloss_per_tM2 = Number(this.sum_totallenthloss_tM2) + Number(this.sum_totallenthloss_per_tM2_first)
      this.sum_number_1_end_tM2 = 0 
      this.sum_number_1_end_tM2 = this.rowexport_type_M2[ax].endless_length_yd * 36
      this.sum_number_2_end_tM2 = 0
      this.sum_number_2_end_tM2 = this.rowexport_type_M2[ax].ttl_marker/this.rowexport_type_M2[ax].length_ydb
      this.sum_plug1_2_end_tM2 = 0
      this.sum_plug1_2_end_tM2 = this.sum_number_1_end_tM2/this.sum_number_2_end_tM2

      this.sum_plug1_2_end_total_tM2 = 0
     
      this.sum_plug1_2_end_total_tM2 = this.sum_plug1_2_end_tM2 * this.rowexport_type_M2[ax].ttl_marker
      this.sum_all_plug1_2_end_total_tM2 = Number(this.sum_all_plug1_2_end_total_tM2) + Number(this.sum_plug1_2_end_total_tM2)

      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_tM2,
        value_table:"M2"
      })
      ws.getCell("N"+countM2).value = Number(this.sum_plug1_2_end_tM2)
      ws.getCell("N"+countM2).numFmt = '0.00';

      countM2++
         }
         if(this.sum_ttl_marker_tM2 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_tM2
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_tM2 = this.summary_plug12_tM2 / this.sum_ttl_marker_tM2
        if(this.sum_plug12_tM2 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_tM2
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_tM2 =  this.summary_widthloss_tM2 / this.sum_ttl_marker_tM2
        if(this.sum_per_width_loss_tM2 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_tM2
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_tM2 =  this.summary_end1_2_tM2/ this.sum_ttl_marker_tM2 //M
      if(this.sum_end1_2_tM2 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_tM2
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_tM2 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_tM2
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_tM2 = this.sum_endless_length_yd_tM2 / this.sum_ttl_marker_tM2 *100 //P
      if(this.sum_avg_end_tM2 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_tM2
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_tM2 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_tM2
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_tM2 = this.sum_spliceloss_tM2 / this.sum_ttl_marker_tM2 *100
      if(this.sum_splicelossper_tM2 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_tM2
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_tM2 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_tM2
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_tM2 = this.sum_totalcutout_tM2 / this.sum_ttl_marker_tM2 *100
      if(this.sum_totalcutoutper_tM2 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_tM2
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_tM2 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_tM2
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_tM2 = this.sum_rement_tM2 / this.sum_ttl_marker_tM2 *100
  if(this.sum_rement_per_tM2 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_tM2
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_tM2
    if(this.sum_totallenthloss_tM2 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_tM2
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspertM2 = this.sum_totallenthlosstM2 / this.sum_ttl_marker_tM2
    if(this.sum_totallenthlosspertM2 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspertM2
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

 ws.getCell("A"+countxM2).value = "Total Table M2"  
      ws.mergeCells("A"+countxM2+":"+"I"+countxM2)
      
      if(this.sum_ttl_marker_tM2 > 0 && isNaN(this.sum_ttl_marker_tM2)==false && isFinite(this.sum_ttl_marker_tM2)==true){
      ws.getCell("J"+countxM2).value = Number(this.sum_ttl_marker_tM2)
      ws.getCell("J"+countxM2).numFmt = '0.00';
      }else{
      ws.getCell("J"+countxM2).value = 0
      ws.getCell("J"+countxM2).numFmt = '0.00';
      }
      
      if(this.sum_plug12_tM2 > 0 && isNaN(this.sum_plug12_tM2)==false && isFinite(this.sum_plug12_tM2)==true){
      ws.getCell("K"+countxM2).value = Number(this.sum_plug12_tM2)
      ws.getCell("K"+countxM2).numFmt = '0.00';
      }else{
      ws.getCell("K"+countxM2).value = 0
      ws.getCell("K"+countxM2).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_tM2 > 0 && isNaN(this.sum_per_width_loss_tM2)==false && isFinite(this.sum_per_width_loss_tM2)==true){
      ws.getCell("L"+countxM2).value = this.sum_per_width_loss_tM2/100
      ws.getCell("L"+countxM2).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countxM2).value = 0.00/100
      ws.getCell("L"+countxM2).numFmt = '0.00%'; 
      }

       this.total_result_plug1_2_tM2 = 0
      this.total_result_plug1_2_tM2 = this.sum_all_plug1_2_end_total_tM2 / this.sum_ttl_marker_tM2

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_tM2
      })
      if(this.total_result_plug1_2_tM2 > 0 && isNaN(this.total_result_plug1_2_tM2)==false && isFinite(this.total_result_plug1_2_tM2)==true){
      ws.getCell("N"+countxM2).value = Number(this.total_result_plug1_2_tM2)
      ws.getCell("N"+countxM2).numFmt = '0.00';
      }else{
      ws.getCell("N"+countxM2).value = 0.00
      ws.getCell("N"+countxM2).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_tM2 > 0 && isNaN(this.sum_endless_length_yd_tM2)==false && isFinite(this.sum_endless_length_yd_tM2)==true){
      ws.getCell("O"+countxM2).value = Number(this.sum_endless_length_yd_tM2)
      ws.getCell("O"+countxM2).numFmt = '0.00';
      }else{
      ws.getCell("O"+countxM2).value = 0.00
      ws.getCell("O"+countxM2).numFmt = '0.00';
      }

      if(this.sum_avg_end_tM2 > 0 && isNaN(this.sum_avg_end_tM2)==false && isFinite(this.sum_avg_end_tM2)==true){
      ws.getCell("P"+countxM2).value = this.sum_avg_end_tM2/100
      ws.getCell("P"+countxM2).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countxM2).value = 0/100
      ws.getCell("P"+countxM2).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_tM2 > 0 && isNaN(this.sum_spliceloss_tM2)==false && isFinite(this.sum_spliceloss_tM2)==true){
      ws.getCell("R"+countxM2).value = Number(this.sum_spliceloss_tM2)
      ws.getCell("R"+countxM2).numFmt = '0.00';
      }else{
       ws.getCell("R"+countxM2).value = 0.00
      ws.getCell("R"+countxM2).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_tM2 > 0 && isNaN(this.sum_splicelossper_tM2)==false && isFinite(this.sum_splicelossper_tM2)==true){
      ws.getCell("S"+countxM2).value = this.sum_splicelossper_tM2/100
      ws.getCell("S"+countxM2).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countxM2).value = 0/100
      ws.getCell("S"+countxM2).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_tM2 > 0 && isNaN(this.sum_totalcutout_tM2)==false && isFinite(this.sum_totalcutout_tM2)==true){
      ws.getCell("U"+countxM2).value = Number(this.sum_totalcutout_tM2)
      ws.getCell("U"+countxM2).numFmt = '0.00';
      }else{
       ws.getCell("U"+countxM2).value = 0.00
      ws.getCell("U"+countxM2).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_tM2 > 0 && isNaN(this.sum_totalcutoutper_tM2)==false && isFinite(this.sum_totalcutoutper_tM2)==true){
      ws.getCell("V"+countxM2).value = this.sum_totalcutoutper_tM2/100
      ws.getCell("V"+countxM2).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countxM2).value = 0.00/100
      ws.getCell("V"+countxM2).numFmt = '0.00%';
      }

       if(this.sum_rement_tM2 > 0 && isNaN(this.sum_rement_tM2)==false && isFinite(this.sum_rement_tM2)==true){
      ws.getCell("X"+countxM2).value = Number(this.sum_rement_tM2)
      ws.getCell("X"+countxM2).numFmt = '0.00';
       }else{
      ws.getCell("X"+countxM2).value = 0.00
      ws.getCell("X"+countxM2).numFmt = '0.00';  
       }

      if(this.sum_rement_per_tM2 > 0 && isNaN(this.sum_rement_per_tM2)==false && isFinite(this.sum_rement_per_tM2)==true){
      ws.getCell("Y"+countxM2).value =  this.sum_rement_per_tM2/100
      ws.getCell("Y"+countxM2).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countxM2).value =  0.00/100
      ws.getCell("Y"+countxM2).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_tM2 > 0 && isNaN(this.sum_totallenthloss_tM2)==false && isFinite(this.sum_totallenthloss_tM2)==true){
      ws.getCell("AA"+countxM2).value = Number(this.sum_totallenthloss_tM2)
      ws.getCell("AA"+countxM2).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countxM2).value = 0
      ws.getCell("AA"+countxM2).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_tM2 = this.sum_totallenthloss_tM2 / this.sum_ttl_marker_tM2 *100
      if(this.last_sum_totallenthloss_per_tM2 > 0 && isNaN(this.last_sum_totallenthloss_per_tM2)==false && isFinite(this.last_sum_totallenthloss_per_tM2)==true){
      ws.getCell("AB"+countxM2).value = this.last_sum_totallenthloss_per_tM2/100
      ws.getCell("AB"+countxM2).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countxM2).value = 0/100
      ws.getCell("AB"+countxM2).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countxM2).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //M3

      if(this.rowexport_type_M2.length == 0){
       
         var countM3 = countM2

       }else{
        var countM3 = countxM2+1
       }
       
      
  
        var countxM3 = countM3 + 1 +Number(this.rowexport_type_M3.length)
      
      this.sum_ttl_marker_tM3 = 0
      this.sum_plug12_tM3 = 0
      this.sum_per_width_loss_tM3 = 0
      this.sum_end1_2_tM3 = 0
      this.sum_endless_length_yd_tM3= 0
      this.sum_avg_end_tM3 = 0
      this.sum_spliceloss_tM3 = 0
      this.sum_splicelossper_tM3 = 0
      this.sum_totalcutout_tM3 = 0
      this.sum_totalcutoutper_tM3 = 0
      this.sum_rement_tM3 = 0
      this.sum_rement_per_tM3 = 0
      this.sum_totallenthloss_tM3 =0 
      this.sum_totallenthlossper_tM3 = 0
      this.summary_widthloss_tM3 = 0
      this.summary_plug12_tM3 = 0
      this.summary_end1_2_tM3 = 0

      
      this.sum_all_plug1_2_end_total_tM3 = 0
     if(this.rowexport_type_M3.length < 1){
      
      count_row_grand++
      
      }else{
      
          for (var ax = 0; ax<this.rowexport_type_M3.length; ax++){ //length +M30
      ws.getCell("A"+countM3).value = this.rowexport_type_M3[ax].mu_date

      ws.getCell("B"+countM3).value = this.rowexport_type_M3[ax].gm_table

      ws.getCell("C"+countM3).value = Number(this.rowexport_type_M3[ax].gm_no_short)
      ws.getCell("C"+countM3).numFmt = '0.00';

      ws.getCell("D"+countM3).value = this.rowexport_type_M3[ax].customer

      ws.getCell("E"+countM3).value = this.rowexport_type_M3[ax].gm_no

      ws.getCell("F"+countM3).value = this.rowexport_type_M3[ax].fabric_type

      ws.getCell("G"+countM3).value = Number(this.rowexport_type_M3[ax].obs_width)
      ws.getCell("G"+countM3).numFmt = '0.00';

      ws.getCell("H"+countM3).value = Number(this.rowexport_type_M3[ax].width_inch)
      ws.getCell("H"+countM3).numFmt = '0.00';


      ws.getCell("I"+countM3).value = Number(this.rowexport_type_M3[ax].length_ydb)
      ws.getCell("I"+countM3).numFmt = '0.00';

      ws.getCell("J"+countM3).value = Number(this.rowexport_type_M3[ax].ttl_marker)
      ws.getCell("J"+countM3).numFmt = '0.00';

      ws.getCell("K"+countM3).value = Number(this.rowexport_type_M3[ax].plug12)
      ws.getCell("K"+countM3).numFmt = '0.00';

      ws.getCell("L"+countM3).value = this.rowexport_type_M3[ax].widthloss/100
      ws.getCell("L"+countM3).numFmt = '0.00%';

      ws.getCell("M"+countM3).value = this.rowexport_type_M3[ax].widthloss_code

      ws.getCell("O"+countM3).value = Number(this.rowexport_type_M3[ax].endless_length_yd)
      ws.getCell("O"+countM3).numFmt = '0.00';

      ws.getCell("P"+countM3).value = this.rowexport_type_M3[ax].avg_end/100
      ws.getCell("P"+countM3).numFmt = '0.00%';

      ws.getCell("Q"+countM3).value = this.rowexport_type_M3[ax].avg_end_code
      

      ws.getCell("R"+countM3).value = Number(this.rowexport_type_M3[ax].spliceloss)
      ws.getCell("R"+countM3).numFmt = '0.00';

      ws.getCell("S"+countM3).value = this.rowexport_type_M3[ax].splicelossper/100
      ws.getCell("S"+countM3).numFmt = '0.00%';

      ws.getCell("T"+countM3).value = this.rowexport_type_M3[ax].per_splice_loss_code
   

      ws.getCell("U"+countM3).value = Number(this.rowexport_type_M3[ax].totalcutout)
      ws.getCell("U"+countM3).numFmt = '0.00';

      ws.getCell("V"+countM3).value = this.rowexport_type_M3[ax].totalcutoutper/100
      ws.getCell("V"+countM3).numFmt = '0.00%';

      ws.getCell("W"+countM3).value = this.rowexport_type_M3[ax].total_cut_out_per_code

      ws.getCell("X"+countM3).value = Number(this.rowexport_type_M3[ax].rement)
      ws.getCell("X"+countM3).numFmt = '0.00';

      ws.getCell("Y"+countM3).value = this.rowexport_type_M3[ax].rement_per/100
      ws.getCell("Y"+countM3).numFmt = '0.00%';

      ws.getCell("Z"+countM3).value = this.rowexport_type_M3[ax].remnants_loss_code

      ws.getCell("AA"+countM3).value = Number(this.rowexport_type_M3[ax].totallenthloss)
      ws.getCell("AA"+countM3).numFmt = '0.00';

      ws.getCell("AB"+countM3).value = this.rowexport_type_M3[ax].totallenthlossper/100
      ws.getCell("AB"+countM3).numFmt = '0.00%';
      
      this.sum_plug12_tM3 = 0
      this.sum_ttl_marker_tM3 = Number(this.sum_ttl_marker_tM3) + Number(this.rowexport_type_M3[ax].ttl_marker)
      this.sum_plug12_tM3 = this.rowexport_type_M3[ax].plug12 * this.rowexport_type_M3[ax].ttl_marker
      this.sum_per_width_loss_tM3 = this.rowexport_type_M3[ax].widthloss * this.rowexport_type_M3[ax].ttl_marker
      this.sum_end1_2_tM3 = this.rowexport_type_M3[ax].end1_2 * this.rowexport_type_M3[ax].ttl_marker
      this.sum_endless_length_yd_tM3 = Number(this.sum_endless_length_yd_tM3) + Number(this.rowexport_type_M3[ax].endless_length_yd)
      this.sum_spliceloss_tM3 = Number(this.sum_spliceloss_tM3) + Number(this.rowexport_type_M3[ax].spliceloss)
      this.sum_totalcutout_tM3 = Number(this.sum_totalcutout_tM3) + Number(this.rowexport_type_M3[ax].totalcutout)
      this.sum_rement_tM3 = Number(this.sum_rement_tM3) + Number(this.rowexport_type_M3[ax].rement)
      this.summary_plug12_tM3 = Number(this.summary_plug12_tM3) + Number(this.sum_plug12_tM3)
      this.summary_widthloss_tM3 = Number(this.summary_widthloss_tM3) + Number(this.sum_per_width_loss_tM3)
      this.summary_end1_2_tM3 = Number(this.summary_end1_2_tM3) + Number(this.sum_end1_2_tM3)
      this.sum_totallenthloss_tM3_first = 0
      this.sum_totallenthloss_tM3_first = Number(this.rowexport_type_M3[ax].endless_length_yd) + Number(this.rowexport_type_M3[ax].spliceloss)
      +Number(this.rowexport_type_M3[ax].totalcutout) + Number(this.rowexport_type_M3[ax].rement)
      this.sum_totallenthloss_tM3 = Number(this.sum_totallenthloss_tM3) + Number(this.sum_totallenthloss_tM3_first)

      this.sum_totallenthloss_per_tM3_first = 0
      this.sum_totallenthloss_per_tM3_first = Number(this.rowexport_type_M3[ax].avg_end) + Number(this.rowexport_type_M3[ax].splicelossper)
      +Number(this.rowexport_type_M3[ax].totalcutoutper) + Number(this.rowexport_type_M3[ax].rement_per)
      this.sum_totallenthloss_per_tM3 = Number(this.sum_totallenthloss_tM3) + Number(this.sum_totallenthloss_per_tM3_first)
      this.sum_number_1_end_tM3 = 0 
      this.sum_number_1_end_tM3 = this.rowexport_type_M3[ax].endless_length_yd * 36
      this.sum_number_2_end_tM3 = 0
      this.sum_number_2_end_tM3 = this.rowexport_type_M3[ax].ttl_marker/this.rowexport_type_M3[ax].length_ydb
      this.sum_plug1_2_end_tM3 = 0
      this.sum_plug1_2_end_tM3 = this.sum_number_1_end_tM3/this.sum_number_2_end_tM3

      this.sum_plug1_2_end_total_tM3 = 0
     
      this.sum_plug1_2_end_total_tM3 = this.sum_plug1_2_end_tM3 * this.rowexport_type_M3[ax].ttl_marker
      this.sum_all_plug1_2_end_total_tM3 = Number(this.sum_all_plug1_2_end_total_tM3) + Number(this.sum_plug1_2_end_total_tM3)

      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_tM3,
        value_table:"M3"
      })     

      ws.getCell("N"+countM3).value = Number(this.sum_plug1_2_end_tM3)
      ws.getCell("N"+countM3).numFmt = '0.00';


      countM3++
         }
         if(this.sum_ttl_marker_tM3 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_tM3
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_tM3 = this.summary_plug12_tM3 / this.sum_ttl_marker_tM3
        if(this.sum_plug12_tM3 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_tM3
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_tM3 =  this.summary_widthloss_tM3 / this.sum_ttl_marker_tM3
        if(this.sum_per_width_loss_tM3 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_tM3
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_tM3 =  this.summary_end1_2_tM3/ this.sum_ttl_marker_tM3 //M
      if(this.sum_end1_2_tM3 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_tM3
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_tM3 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_tM3
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_tM3 = this.sum_endless_length_yd_tM3 / this.sum_ttl_marker_tM3 *100 //P
      if(this.sum_avg_end_tM3 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_tM3
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_tM3 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_tM3
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_tM3 = this.sum_spliceloss_tM3 / this.sum_ttl_marker_tM3 *100
      if(this.sum_splicelossper_tM3 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_tM3
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_tM3 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_tM3
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_tM3 = this.sum_totalcutout_tM3 / this.sum_ttl_marker_tM3 *100
      if(this.sum_totalcutoutper_tM3 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_tM3
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_tM3 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_tM3
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_tM3 = this.sum_rement_tM3 / this.sum_ttl_marker_tM3 *100
  if(this.sum_rement_per_tM3 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_tM3
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_tM3
    if(this.sum_totallenthloss_tM3 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_tM3
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspertM3 = this.sum_totallenthlosstM3 / this.sum_ttl_marker_tM3
    if(this.sum_totallenthlosspertM3 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspertM3
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

ws.getCell("A"+countxM3).value = "Total Table M3"  
      ws.mergeCells("A"+countxM3+":"+"I"+countxM3)
      
      if(this.sum_ttl_marker_tM3 > 0 && isNaN(this.sum_ttl_marker_tM3)==false && isFinite(this.sum_ttl_marker_tM3)==true){
      ws.getCell("J"+countxM3).value = Number(this.sum_ttl_marker_tM3)
      ws.getCell("J"+countxM3).numFmt = '0.00';
      }else{
      ws.getCell("J"+countxM3).value = 0
      ws.getCell("J"+countxM3).numFmt = '0.00';
      }
      console.log(this.sum_plug12_tM3)
      if(this.sum_plug12_tM3 > 0 && isNaN(this.sum_plug12_tM3)==false && isFinite(this.sum_plug12_tM3)==true){
      ws.getCell("K"+countxM3).value = Number(this.sum_plug12_tM3)
      ws.getCell("K"+countxM3).numFmt = '0.00';
      }else{
      ws.getCell("K"+countxM3).value = 0
      ws.getCell("K"+countxM3).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_tM3 > 0 && isNaN(this.sum_per_width_loss_tM3)==false && isFinite(this.sum_per_width_loss_tM3)==true){
      ws.getCell("L"+countxM3).value = this.sum_per_width_loss_tM3/100
      ws.getCell("L"+countxM3).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countxM3).value = 0.00/100
      ws.getCell("L"+countxM3).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_tM3 = 0
      this.total_result_plug1_2_tM3 = this.sum_all_plug1_2_end_total_tM3 / this.sum_ttl_marker_tM3

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_tM3
      })
      if(this.total_result_plug1_2_tM3 > 0 && isNaN(this.total_result_plug1_2_tM3)==false && isFinite(this.total_result_plug1_2_tM3)==true){
      ws.getCell("N"+countxM3).value = Number(this.total_result_plug1_2_tM3)
      ws.getCell("N"+countxM3).numFmt = '0.00';
      }else{
      ws.getCell("N"+countxM3).value = 0.00
      ws.getCell("N"+countxM3).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_tM3 > 0 && isNaN(this.sum_endless_length_yd_tM3)==false && isFinite(this.sum_endless_length_yd_tM3)==true){
      ws.getCell("O"+countxM3).value = Number(this.sum_endless_length_yd_tM3)
      ws.getCell("O"+countxM3).numFmt = '0.00';
      }else{
      ws.getCell("O"+countxM3).value = 0.00
      ws.getCell("O"+countxM3).numFmt = '0.00';
      }

      if(this.sum_avg_end_tM3 > 0 && isNaN(this.sum_avg_end_tM3)==false && isFinite(this.sum_avg_end_tM3)==true){
      ws.getCell("P"+countxM3).value = this.sum_avg_end_tM3/100
      ws.getCell("P"+countxM3).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countxM3).value = 0/100
      ws.getCell("P"+countxM3).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_tM3 > 0 && isNaN(this.sum_spliceloss_tM3)==false && isFinite(this.sum_spliceloss_tM3)==true){
      ws.getCell("R"+countxM3).value = Number(this.sum_spliceloss_tM3)
      ws.getCell("R"+countxM3).numFmt = '0.00';
      }else{
       ws.getCell("R"+countxM3).value = 0.00
      ws.getCell("R"+countxM3).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_tM3 > 0 && isNaN(this.sum_splicelossper_tM3)==false && isFinite(this.sum_splicelossper_tM3)==true){
      ws.getCell("S"+countxM3).value = this.sum_splicelossper_tM3/100
      ws.getCell("S"+countxM3).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countxM3).value = 0/100
      ws.getCell("S"+countxM3).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_tM3 > 0 && isNaN(this.sum_totalcutout_tM3)==false && isFinite(this.sum_totalcutout_tM3)==true){
      ws.getCell("U"+countxM3).value = Number(this.sum_totalcutout_tM3)
      ws.getCell("U"+countxM3).numFmt = '0.00';
      }else{
       ws.getCell("U"+countxM3).value = 0.00
      ws.getCell("U"+countxM3).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_tM3 > 0 && isNaN(this.sum_totalcutoutper_tM3)==false && isFinite(this.sum_totalcutoutper_tM3)==true){
      ws.getCell("V"+countxM3).value = this.sum_totalcutoutper_tM3/100
      ws.getCell("V"+countxM3).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countxM3).value = 0.00/100
      ws.getCell("V"+countxM3).numFmt = '0.00%';
      }

       if(this.sum_rement_tM3 > 0 && isNaN(this.sum_rement_tM3)==false && isFinite(this.sum_rement_tM3)==true){
      ws.getCell("X"+countxM3).value = Number(this.sum_rement_tM3)
      ws.getCell("X"+countxM3).numFmt = '0.00';
       }else{
      ws.getCell("X"+countxM3).value = 0.00
      ws.getCell("X"+countxM3).numFmt = '0.00';  
       }

      if(this.sum_rement_per_tM3 > 0 && isNaN(this.sum_rement_per_tM3)==false && isFinite(this.sum_rement_per_tM3)==true){
      ws.getCell("Y"+countxM3).value =  this.sum_rement_per_tM3/100
      ws.getCell("Y"+countxM3).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countxM3).value =  0.00/100
      ws.getCell("Y"+countxM3).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_tM3 > 0 && isNaN(this.sum_totallenthloss_tM3)==false && isFinite(this.sum_totallenthloss_tM3)==true){
      ws.getCell("AA"+countxM3).value = Number(this.sum_totallenthloss_tM3)
      ws.getCell("AA"+countxM3).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countxM3).value = 0
      ws.getCell("AA"+countxM3).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_tM3 = this.sum_totallenthloss_tM3 / this.sum_ttl_marker_tM3 *100
      if(this.last_sum_totallenthloss_per_tM3 > 0 && isNaN(this.last_sum_totallenthloss_per_tM3)==false && isFinite(this.last_sum_totallenthloss_per_tM3)==true){
      ws.getCell("AB"+countxM3).value = this.last_sum_totallenthloss_per_tM3/100
      ws.getCell("AB"+countxM3).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countxM3).value = 0/100
      ws.getCell("AB"+countxM3).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countxM3).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //M4

      if(this.rowexport_type_M3.length == 0){
       
         var countM4 = countM3

       }else{
        var countM4 = countxM3+1
       }
       
  
        var countxM4 = countM4 + 1 +Number(this.rowexport_type_M4.length)
      
      this.sum_ttl_marker_tM4 = 0
      this.sum_plug12_tM4 = 0
      this.sum_per_width_loss_tM4 = 0
      this.sum_end1_2_tM4 = 0
      this.sum_endless_length_yd_tM4= 0
      this.sum_avg_end_tM4 = 0
      this.sum_spliceloss_tM4 = 0
      this.sum_splicelossper_tM4 = 0
      this.sum_totalcutout_tM4 = 0
      this.sum_totalcutoutper_tM4 = 0
      this.sum_rement_tM4 = 0
      this.sum_rement_per_tM4 = 0
      this.sum_totallenthloss_tM4 =0 
      this.sum_totallenthlossper_tM4 = 0
      this.summary_widthloss_tM4 = 0
      this.summary_plug12_tM4 = 0
      this.summary_end1_2_tM4 = 0

      
     this.sum_all_plug1_2_end_total_tM4 = 0
      if(this.rowexport_type_M4.length < 1){
      count_row_grand++
      
      
      }else{
          for (var ax = 0; ax<this.rowexport_type_M4.length; ax++){ //length +M40
      ws.getCell("A"+countM4).value = this.rowexport_type_M4[ax].mu_date

      ws.getCell("B"+countM4).value = this.rowexport_type_M4[ax].gm_table

      ws.getCell("C"+countM4).value = Number(this.rowexport_type_M4[ax].gm_no_short)
      ws.getCell("C"+countM4).numFmt = '0.00';

      ws.getCell("D"+countM4).value = this.rowexport_type_M4[ax].customer

      ws.getCell("E"+countM4).value = this.rowexport_type_M4[ax].gm_no

      ws.getCell("F"+countM4).value = this.rowexport_type_M4[ax].fabric_type

      ws.getCell("G"+countM4).value = Number(this.rowexport_type_M4[ax].obs_width)
      ws.getCell("G"+countM4).numFmt = '0.00';

      ws.getCell("H"+countM4).value = Number(this.rowexport_type_M4[ax].width_inch)
      ws.getCell("H"+countM4).numFmt = '0.00';


      ws.getCell("I"+countM4).value = Number(this.rowexport_type_M4[ax].length_ydb)
      ws.getCell("I"+countM4).numFmt = '0.00';

      ws.getCell("J"+countM4).value = Number(this.rowexport_type_M4[ax].ttl_marker)
      ws.getCell("J"+countM4).numFmt = '0.00';

      ws.getCell("K"+countM4).value = Number(this.rowexport_type_M4[ax].plug12)
      ws.getCell("K"+countM4).numFmt = '0.00';

      ws.getCell("L"+countM4).value = this.rowexport_type_M4[ax].widthloss/100
      ws.getCell("L"+countM4).numFmt = '0.00%';

      ws.getCell("M"+countM4).value = this.rowexport_type_M4[ax].widthloss_code

      ws.getCell("O"+countM4).value = Number(this.rowexport_type_M4[ax].endless_length_yd)
      ws.getCell("O"+countM4).numFmt = '0.00';

      ws.getCell("P"+countM4).value = this.rowexport_type_M4[ax].avg_end/100
      ws.getCell("P"+countM4).numFmt = '0.00%';

      ws.getCell("Q"+countM4).value = this.rowexport_type_M4[ax].avg_end_code
      

      ws.getCell("R"+countM4).value = Number(this.rowexport_type_M4[ax].spliceloss)
      ws.getCell("R"+countM4).numFmt = '0.00';

      ws.getCell("S"+countM4).value = this.rowexport_type_M4[ax].splicelossper/100
      ws.getCell("S"+countM4).numFmt = '0.00%';

      ws.getCell("T"+countM4).value = this.rowexport_type_M4[ax].per_splice_loss_code
   

      ws.getCell("U"+countM4).value = Number(this.rowexport_type_M4[ax].totalcutout)
      ws.getCell("U"+countM4).numFmt = '0.00';

      ws.getCell("V"+countM4).value = this.rowexport_type_M4[ax].totalcutoutper/100
      ws.getCell("V"+countM4).numFmt = '0.00%';

      ws.getCell("W"+countM4).value = this.rowexport_type_M4[ax].total_cut_out_per_code

      ws.getCell("X"+countM4).value = Number(this.rowexport_type_M4[ax].rement)
      ws.getCell("X"+countM4).numFmt = '0.00';

      ws.getCell("Y"+countM4).value = this.rowexport_type_M4[ax].rement_per/100
      ws.getCell("Y"+countM4).numFmt = '0.00%';

      ws.getCell("Z"+countM4).value = this.rowexport_type_M4[ax].remnants_loss_code

      ws.getCell("AA"+countM4).value = Number(this.rowexport_type_M4[ax].totallenthloss)
      ws.getCell("AA"+countM4).numFmt = '0.00';

      ws.getCell("AB"+countM4).value = this.rowexport_type_M4[ax].totallenthlossper/100
      ws.getCell("AB"+countM4).numFmt = '0.00%';
      
      this.sum_plug12_tM4 = 0
      this.sum_ttl_marker_tM4 = Number(this.sum_ttl_marker_tM4) + Number(this.rowexport_type_M4[ax].ttl_marker)
      this.sum_plug12_tM4 = this.rowexport_type_M4[ax].plug12 * this.rowexport_type_M4[ax].ttl_marker
      this.sum_per_width_loss_tM4 = this.rowexport_type_M4[ax].widthloss * this.rowexport_type_M4[ax].ttl_marker
      this.sum_end1_2_tM4 = this.rowexport_type_M4[ax].end1_2 * this.rowexport_type_M4[ax].ttl_marker
      this.sum_endless_length_yd_tM4 = Number(this.sum_endless_length_yd_tM4) + Number(this.rowexport_type_M4[ax].endless_length_yd)
      this.sum_spliceloss_tM4 = Number(this.sum_spliceloss_tM4) + Number(this.rowexport_type_M4[ax].spliceloss)
      this.sum_totalcutout_tM4 = Number(this.sum_totalcutout_tM4) + Number(this.rowexport_type_M4[ax].totalcutout)
      this.sum_rement_tM4 = Number(this.sum_rement_tM4) + Number(this.rowexport_type_M4[ax].rement)
      this.summary_plug12_tM4 = Number(this.summary_plug12_tM4) + Number(this.sum_plug12_tM4)
      this.summary_widthloss_tM4 = Number(this.summary_widthloss_tM4) + Number(this.sum_per_width_loss_tM4)
      this.summary_end1_2_tM4 = Number(this.summary_end1_2_tM4) + Number(this.sum_end1_2_tM4)
      this.sum_totallenthloss_tM4_first = 0
      this.sum_totallenthloss_tM4_first = Number(this.rowexport_type_M4[ax].endless_length_yd) + Number(this.rowexport_type_M4[ax].spliceloss)
      +Number(this.rowexport_type_M4[ax].totalcutout) + Number(this.rowexport_type_M4[ax].rement)
      this.sum_totallenthloss_tM4 = Number(this.sum_totallenthloss_tM4) + Number(this.sum_totallenthloss_tM4_first)

      this.sum_totallenthloss_per_tM4_first = 0
      this.sum_totallenthloss_per_tM4_first = Number(this.rowexport_type_M4[ax].avg_end) + Number(this.rowexport_type_M4[ax].splicelossper)
      +Number(this.rowexport_type_M4[ax].totalcutoutper) + Number(this.rowexport_type_M4[ax].rement_per)
      this.sum_totallenthloss_per_tM4 = Number(this.sum_totallenthloss_tM4) + Number(this.sum_totallenthloss_per_tM4_first)
      this.sum_number_1_end_tM4 = 0 
      this.sum_number_1_end_tM4 = this.rowexport_type_M4[ax].endless_length_yd * 36
      this.sum_number_2_end_tM4 = 0
      this.sum_number_2_end_tM4 = this.rowexport_type_M4[ax].ttl_marker/this.rowexport_type_M4[ax].length_ydb
      this.sum_plug1_2_end_tM4 = 0
      this.sum_plug1_2_end_tM4 = this.sum_number_1_end_tM4/this.sum_number_2_end_tM4

      this.sum_plug1_2_end_total_tM4 = 0
     
      this.sum_plug1_2_end_total_tM4 = this.sum_plug1_2_end_tM4 * this.rowexport_type_M4[ax].ttl_marker
      this.sum_all_plug1_2_end_total_tM4 = Number(this.sum_all_plug1_2_end_total_tM4) + Number(this.sum_plug1_2_end_total_tM4)

      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_tM4,
        value_table:"M4"
      })

      ws.getCell("N"+countM4).value = Number(this.sum_plug1_2_end_tM4)
      ws.getCell("N"+countM4).numFmt = '0.00';


      countM4++
         }
         if(this.sum_ttl_marker_tM4 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_tM4
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_tM4 = this.summary_plug12_tM4 / this.sum_ttl_marker_tM4
        if(this.sum_plug12_tM4 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_tM4
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_tM4 =  this.summary_widthloss_tM4 / this.sum_ttl_marker_tM4
        if(this.sum_per_width_loss_tM4 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_tM4
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_tM4 =  this.summary_end1_2_tM4/ this.sum_ttl_marker_tM4 //M
      if(this.sum_end1_2_tM4 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_tM4
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_tM4 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_tM4
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_tM4 = this.sum_endless_length_yd_tM4 / this.sum_ttl_marker_tM4 *100 //P
      if(this.sum_avg_end_tM4 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_tM4
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_tM4 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_tM4
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_tM4 = this.sum_spliceloss_tM4 / this.sum_ttl_marker_tM4 *100
      if(this.sum_splicelossper_tM4 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_tM4
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_tM4 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_tM4
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_tM4 = this.sum_totalcutout_tM4 / this.sum_ttl_marker_tM4 *100
      if(this.sum_totalcutoutper_tM4 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_tM4
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_tM4 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_tM4
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_tM4 = this.sum_rement_tM4 / this.sum_ttl_marker_tM4 *100
  if(this.sum_rement_per_tM4 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_tM4
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_tM4
    if(this.sum_totallenthloss_tM4 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_tM4
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspertM4 = this.sum_totallenthlosstM4 / this.sum_ttl_marker_tM4
    if(this.sum_totallenthlosspertM4 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspertM4
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }

 ws.getCell("A"+countxM4).value = "Total Table M4"  
      ws.mergeCells("A"+countxM4+":"+"I"+countxM4)
      
      if(this.sum_ttl_marker_tM4 > 0 && isNaN(this.sum_ttl_marker_tM4)==false && isFinite(this.sum_ttl_marker_tM4)==true){
      ws.getCell("J"+countxM4).value = Number(this.sum_ttl_marker_tM4)
      ws.getCell("J"+countxM4).numFmt = '0.00';
      }else{
      ws.getCell("J"+countxM4).value = 0
      ws.getCell("J"+countxM4).numFmt = '0.00';
      }

      if(this.sum_plug12_tM4 > 0 && isNaN(this.sum_plug12_tM4)==false && isFinite(this.sum_plug12_tM4)==true){
      ws.getCell("K"+countxM4).value = Number(this.sum_plug12_tM4)
      ws.getCell("K"+countxM4).numFmt = '0.00';
      }else{
      ws.getCell("K"+countxM4).value = 0
      ws.getCell("K"+countxM4).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_tM4 > 0 && isNaN(this.sum_per_width_loss_tM4)==false && isFinite(this.sum_per_width_loss_tM4)==true){
      ws.getCell("L"+countxM4).value = this.sum_per_width_loss_tM4/100
      ws.getCell("L"+countxM4).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countxM4).value = 0.00/100
      ws.getCell("L"+countxM4).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_tM4 = 0
      this.total_result_plug1_2_tM4 = this.sum_all_plug1_2_end_total_tM4 / this.sum_ttl_marker_tM4

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_tM4
      })
      if(this.total_result_plug1_2_tM4 > 0 && isNaN(this.total_result_plug1_2_tM4)==false && isFinite(this.total_result_plug1_2_tM4)==true){
      ws.getCell("N"+countxM4).value = Number(this.total_result_plug1_2_tM4)
      ws.getCell("N"+countxM4).numFmt = '0.00';
      }else{
      ws.getCell("N"+countxM4).value = 0.00
      ws.getCell("N"+countxM4).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_tM4 > 0 && isNaN(this.sum_endless_length_yd_tM4)==false && isFinite(this.sum_endless_length_yd_tM4)==true){
      ws.getCell("O"+countxM4).value = Number(this.sum_endless_length_yd_tM4)
      ws.getCell("O"+countxM4).numFmt = '0.00';
      }else{
      ws.getCell("O"+countxM4).value = 0.00
      ws.getCell("O"+countxM4).numFmt = '0.00';
      }

      if(this.sum_avg_end_tM4 > 0 && isNaN(this.sum_avg_end_tM4)==false && isFinite(this.sum_avg_end_tM4)==true){
      ws.getCell("P"+countxM4).value = this.sum_avg_end_tM4/100
      ws.getCell("P"+countxM4).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countxM4).value = 0/100
      ws.getCell("P"+countxM4).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_tM4 > 0 && isNaN(this.sum_spliceloss_tM4)==false && isFinite(this.sum_spliceloss_tM4)==true){
      ws.getCell("R"+countxM4).value = Number(this.sum_spliceloss_tM4)
      ws.getCell("R"+countxM4).numFmt = '0.00';
      }else{
       ws.getCell("R"+countxM4).value = 0.00
      ws.getCell("R"+countxM4).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_tM4 > 0 && isNaN(this.sum_splicelossper_tM4)==false && isFinite(this.sum_splicelossper_tM4)==true){
      ws.getCell("S"+countxM4).value = this.sum_splicelossper_tM4/100
      ws.getCell("S"+countxM4).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countxM4).value = 0/100
      ws.getCell("S"+countxM4).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_tM4 > 0 && isNaN(this.sum_totalcutout_tM4)==false && isFinite(this.sum_totalcutout_tM4)==true){
      ws.getCell("U"+countxM4).value = Number(this.sum_totalcutout_tM4)
      ws.getCell("U"+countxM4).numFmt = '0.00';
      }else{
       ws.getCell("U"+countxM4).value = 0.00
      ws.getCell("U"+countxM4).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_tM4 > 0 && isNaN(this.sum_totalcutoutper_tM4)==false && isFinite(this.sum_totalcutoutper_tM4)==true){
      ws.getCell("V"+countxM4).value = this.sum_totalcutoutper_tM4/100
      ws.getCell("V"+countxM4).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countxM4).value = 0.00/100
      ws.getCell("V"+countxM4).numFmt = '0.00%';
      }

       if(this.sum_rement_tM4 > 0 && isNaN(this.sum_rement_tM4)==false && isFinite(this.sum_rement_tM4)==true){
      ws.getCell("X"+countxM4).value = Number(this.sum_rement_tM4)
      ws.getCell("X"+countxM4).numFmt = '0.00';
       }else{
      ws.getCell("X"+countxM4).value = 0.00
      ws.getCell("X"+countxM4).numFmt = '0.00';  
       }

      if(this.sum_rement_per_tM4 > 0 && isNaN(this.sum_rement_per_tM4)==false && isFinite(this.sum_rement_per_tM4)==true){
      ws.getCell("Y"+countxM4).value =  this.sum_rement_per_tM4/100
      ws.getCell("Y"+countxM4).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countxM4).value =  0.00/100
      ws.getCell("Y"+countxM4).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_tM4 > 0 && isNaN(this.sum_totallenthloss_tM4)==false && isFinite(this.sum_totallenthloss_tM4)==true){
      ws.getCell("AA"+countxM4).value = Number(this.sum_totallenthloss_tM4)
      ws.getCell("AA"+countxM4).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countxM4).value = 0
      ws.getCell("AA"+countxM4).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_tM4 = this.sum_totallenthloss_tM4 / this.sum_ttl_marker_tM4 *100
      if(this.last_sum_totallenthloss_per_tM4 > 0 && isNaN(this.last_sum_totallenthloss_per_tM4)==false && isFinite(this.last_sum_totallenthloss_per_tM4)==true){
      ws.getCell("AB"+countxM4).value = this.last_sum_totallenthloss_per_tM4/100
      ws.getCell("AB"+countxM4).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countxM4).value = 0/100
      ws.getCell("AB"+countxM4).numFmt = '0.00%'; 
      }

      
    

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countxM4).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //A5
      if(this.rowexport_type_M4.length == 0){
       
         var countM5 = countM4

       }else{
         var countM5 = countxM4+1
       }
       
      
  
        var countxM5 = countM5 + 1 +Number(this.rowexport_type_M5.length)
      
      this.sum_ttl_marker_tM5 = 0
      this.sum_plug12_tM5 = 0
      this.sum_per_width_loss_tM5 = 0
      this.sum_end1_2_tM5 = 0
      this.sum_endless_length_yd_tM5= 0
      this.sum_avg_end_tM5 = 0
      this.sum_spliceloss_tM5 = 0
      this.sum_splicelossper_tM5 = 0
      this.sum_totalcutout_tM5 = 0
      this.sum_totalcutoutper_tM5 = 0
      this.sum_rement_tM5 = 0
      this.sum_rement_per_tM5 = 0
      this.sum_totallenthloss_tM5 =0 
      this.sum_totallenthlossper_tM5 = 0
      this.summary_widthloss_tM5 = 0
      this.summary_plug12_tM5 = 0
      this.summary_end1_2_tM5 = 0

      
     this.sum_all_plug1_2_end_total_tM5 = 0
      if(this.rowexport_type_M5.length < 1){
      
      count_row_grand++
      
      }else{
          for (var ax = 0; ax<this.rowexport_type_M5.length; ax++){ //length +M50
      ws.getCell("A"+countM5).value = this.rowexport_type_M5[ax].mu_date

      ws.getCell("B"+countM5).value = this.rowexport_type_M5[ax].gm_table

      ws.getCell("C"+countM5).value = Number(this.rowexport_type_M5[ax].gm_no_short)
      ws.getCell("C"+countM5).numFmt = '0.00';

      ws.getCell("D"+countM5).value = this.rowexport_type_M5[ax].customer

      ws.getCell("E"+countM5).value = this.rowexport_type_M5[ax].gm_no

      ws.getCell("F"+countM5).value = this.rowexport_type_M5[ax].fabric_type

      ws.getCell("G"+countM5).value = Number(this.rowexport_type_M5[ax].obs_width)
      ws.getCell("G"+countM5).numFmt = '0.00';

      ws.getCell("H"+countM5).value = Number(this.rowexport_type_M5[ax].width_inch)
      ws.getCell("H"+countM5).numFmt = '0.00';


      ws.getCell("I"+countM5).value = Number(this.rowexport_type_M5[ax].length_ydb)
      ws.getCell("I"+countM5).numFmt = '0.00';

      ws.getCell("J"+countM5).value = Number(this.rowexport_type_M5[ax].ttl_marker)
      ws.getCell("J"+countM5).numFmt = '0.00';

      ws.getCell("K"+countM5).value = Number(this.rowexport_type_M5[ax].plug12)
      ws.getCell("K"+countM5).numFmt = '0.00';

      ws.getCell("L"+countM5).value = this.rowexport_type_M5[ax].widthloss/100
      ws.getCell("L"+countM5).numFmt = '0.00%';

      ws.getCell("M"+countM5).value = this.rowexport_type_M5[ax].widthloss_code


      ws.getCell("O"+countM5).value = Number(this.rowexport_type_M5[ax].endless_length_yd)
      ws.getCell("O"+countM5).numFmt = '0.00';

      ws.getCell("P"+countM5).value = this.rowexport_type_M5[ax].avg_end/100
      ws.getCell("P"+countM5).numFmt = '0.00%';

      ws.getCell("Q"+countM5).value = this.rowexport_type_M5[ax].avg_end_code
      

      ws.getCell("R"+countM5).value = Number(this.rowexport_type_M5[ax].spliceloss)
      ws.getCell("R"+countM5).numFmt = '0.00';

      ws.getCell("S"+countM5).value = this.rowexport_type_M5[ax].splicelossper/100
      ws.getCell("S"+countM5).numFmt = '0.00%';

      ws.getCell("T"+countM5).value = this.rowexport_type_M5[ax].per_splice_loss_code
   

      ws.getCell("U"+countM5).value = Number(this.rowexport_type_M5[ax].totalcutout)
      ws.getCell("U"+countM5).numFmt = '0.00';

      ws.getCell("V"+countM5).value = this.rowexport_type_M5[ax].totalcutoutper/100
      ws.getCell("V"+countM5).numFmt = '0.00%';

      ws.getCell("W"+countM5).value = this.rowexport_type_M5[ax].total_cut_out_per_code

      ws.getCell("X"+countM5).value = Number(this.rowexport_type_M5[ax].rement)
      ws.getCell("X"+countM5).numFmt = '0.00';

      ws.getCell("Y"+countM5).value = this.rowexport_type_M5[ax].rement_per/100
      ws.getCell("Y"+countM5).numFmt = '0.00%';

      ws.getCell("Z"+countM5).value = this.rowexport_type_M5[ax].remnants_loss_code

      ws.getCell("AA"+countM5).value = Number(this.rowexport_type_M5[ax].totallenthloss)
      ws.getCell("AA"+countM5).numFmt = '0.00';

      ws.getCell("AB"+countM5).value = this.rowexport_type_M5[ax].totallenthlossper/100
      ws.getCell("AB"+countM5).numFmt = '0.00%';
      
      this.sum_plug12_tM5 = 0
      this.sum_ttl_marker_tM5 = Number(this.sum_ttl_marker_tM5) + Number(this.rowexport_type_M5[ax].ttl_marker)
      this.sum_plug12_tM5 = this.rowexport_type_M5[ax].plug12 * this.rowexport_type_M5[ax].ttl_marker
      this.sum_per_width_loss_tM5 = this.rowexport_type_M5[ax].widthloss * this.rowexport_type_M5[ax].ttl_marker
      this.sum_end1_2_tM5 = this.rowexport_type_M5[ax].end1_2 * this.rowexport_type_M5[ax].ttl_marker
      this.sum_endless_length_yd_tM5 = Number(this.sum_endless_length_yd_tM5) + Number(this.rowexport_type_M5[ax].endless_length_yd)
      this.sum_spliceloss_tM5 = Number(this.sum_spliceloss_tM5) + Number(this.rowexport_type_M5[ax].spliceloss)
      this.sum_totalcutout_tM5 = Number(this.sum_totalcutout_tM5) + Number(this.rowexport_type_M5[ax].totalcutout)
      this.sum_rement_tM5 = Number(this.sum_rement_tM5) + Number(this.rowexport_type_M5[ax].rement)
      this.summary_plug12_tM5 = Number(this.summary_plug12_tM5) + Number(this.sum_plug12_tM5)
      this.summary_widthloss_tM5 = Number(this.summary_widthloss_tM5) + Number(this.sum_per_width_loss_tM5)
      this.summary_end1_2_tM5 = Number(this.summary_end1_2_tM5) + Number(this.sum_end1_2_tM5)
      this.sum_totallenthloss_tM5_first = 0
      this.sum_totallenthloss_tM5_first = Number(this.rowexport_type_M5[ax].endless_length_yd) + Number(this.rowexport_type_M5[ax].spliceloss)
      +Number(this.rowexport_type_M5[ax].totalcutout) + Number(this.rowexport_type_M5[ax].rement)
      this.sum_totallenthloss_tM5 = Number(this.sum_totallenthloss_tM5) + Number(this.sum_totallenthloss_tM5_first)

      this.sum_totallenthloss_per_tM5_first = 0
      this.sum_totallenthloss_per_tM5_first = Number(this.rowexport_type_M5[ax].avg_end) + Number(this.rowexport_type_M5[ax].splicelossper)
      +Number(this.rowexport_type_M5[ax].totalcutoutper) + Number(this.rowexport_type_M5[ax].rement_per)
      this.sum_totallenthloss_per_tM5 = Number(this.sum_totallenthloss_tM5) + Number(this.sum_totallenthloss_per_tM5_first)
      this.sum_number_1_end_tM5 = 0 
      this.sum_number_1_end_tM5 = this.rowexport_type_M5[ax].endless_length_yd * 36
      this.sum_number_2_end_tM5 = 0
      this.sum_number_2_end_tM5 = this.rowexport_type_M5[ax].ttl_marker/this.rowexport_type_M5[ax].length_ydb
      this.sum_plug1_2_end_tM5 = 0
      this.sum_plug1_2_end_tM5 = this.sum_number_1_end_tM5/this.sum_number_2_end_tM5

      this.sum_plug1_2_end_total_tM5 = 0
     
      this.sum_plug1_2_end_total_tM5 = this.sum_plug1_2_end_tM5 * this.rowexport_type_M5[ax].ttl_marker
      this.sum_all_plug1_2_end_total_tM5 = Number(this.sum_all_plug1_2_end_total_tM5) + Number(this.sum_plug1_2_end_total_tM5)
          
      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_tM5,
        value_table:"M5"
      })
      ws.getCell("N"+countM5).value = Number(this.sum_plug1_2_end_tM5)
      ws.getCell("N"+countM5).numFmt = '0.00';
      countM5++
         }
         if(this.sum_ttl_marker_tM5 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_tM5
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_tM5 = this.summary_plug12_tM5 / this.sum_ttl_marker_tM5
        if(this.sum_plug12_tM5 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_tM5
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_tM5 =  this.summary_widthloss_tM5 / this.sum_ttl_marker_tM5
        if(this.sum_per_width_loss_tM5 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_tM5
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_tM5 =  this.summary_end1_2_tM5/ this.sum_ttl_marker_tM5 //M
      if(this.sum_end1_2_tM5 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_tM5
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_tM5 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_tM5
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_tM5 = this.sum_endless_length_yd_tM5 / this.sum_ttl_marker_tM5 *100 //P
      if(this.sum_avg_end_tM5 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_tM5
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_tM5 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_tM5
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_tM5 = this.sum_spliceloss_tM5 / this.sum_ttl_marker_tM5 *100
      if(this.sum_splicelossper_tM5 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_tM5
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_tM5 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_tM5
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_tM5 = this.sum_totalcutout_tM5 / this.sum_ttl_marker_tM5 *100
      if(this.sum_totalcutoutper_tM5 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_tM5
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_tM5 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_tM5
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_tM5 = this.sum_rement_tM5 / this.sum_ttl_marker_tM5 *100
  if(this.sum_rement_per_tM5 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_tM5
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_tM5
    if(this.sum_totallenthloss_tM5 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_tM5
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspertM5 = this.sum_totallenthlosstM5 / this.sum_ttl_marker_tM5
    if(this.sum_totallenthlosspertM5 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspertM5
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }
    
ws.getCell("A"+countxM5).value = "Total Table M5"  
      ws.mergeCells("A"+countxM5+":"+"I"+countxM5)
      
      if(this.sum_ttl_marker_tM5 > 0 && isNaN(this.sum_ttl_marker_tM5)==false && isFinite(this.sum_ttl_marker_tM5)==true){
      ws.getCell("J"+countxM5).value = Number(this.sum_ttl_marker_tM5)
      ws.getCell("J"+countxM5).numFmt = '0.00';
      }else{
      ws.getCell("J"+countxM5).value = 0
      ws.getCell("J"+countxM5).numFmt = '0.00';
      }

      if(this.sum_plug12_tM5 > 0 && isNaN(this.sum_plug12_tM5)==false && isFinite(this.sum_plug12_tM5)==true){
      ws.getCell("K"+countxM5).value = Number(this.sum_plug12_tM5)
      ws.getCell("K"+countxM5).numFmt = '0.00';
      }else{
      ws.getCell("K"+countxM5).value = 0
      ws.getCell("K"+countxM5).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_tM5 > 0 && isNaN(this.sum_per_width_loss_tM5)==false && isFinite(this.sum_per_width_loss_tM5)==true){
      ws.getCell("L"+countxM5).value = this.sum_per_width_loss_tM5/100
      ws.getCell("L"+countxM5).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countxM5).value = 0.00/100
      ws.getCell("L"+countxM5).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_tM5 = 0
      this.total_result_plug1_2_tM5 = this.sum_all_plug1_2_end_total_tM5 / this.sum_ttl_marker_tM5

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_tM5
      })
      if(this.total_result_plug1_2_tM5 > 0 && isNaN(this.total_result_plug1_2_tM5)==false && isFinite(this.total_result_plug1_2_tM5)==true){
      ws.getCell("N"+countxM5).value = Number(this.total_result_plug1_2_tM5)
      ws.getCell("N"+countxM5).numFmt = '0.00';
      }else{
      ws.getCell("N"+countxM5).value = 0.00
      ws.getCell("N"+countxM5).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_tM5 > 0 && isNaN(this.sum_endless_length_yd_tM5)==false && isFinite(this.sum_endless_length_yd_tM5)==true){
      ws.getCell("O"+countxM5).value = Number(this.sum_endless_length_yd_tM5)
      ws.getCell("O"+countxM5).numFmt = '0.00';
      }else{
      ws.getCell("O"+countxM5).value = 0.00
      ws.getCell("O"+countxM5).numFmt = '0.00';
      }

      if(this.sum_avg_end_tM5 > 0 && isNaN(this.sum_avg_end_tM5)==false && isFinite(this.sum_avg_end_tM5)==true){
      ws.getCell("P"+countxM5).value = this.sum_avg_end_tM5/100
      ws.getCell("P"+countxM5).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countxM5).value = 0/100
      ws.getCell("P"+countxM5).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_tM5 > 0 && isNaN(this.sum_spliceloss_tM5)==false && isFinite(this.sum_spliceloss_tM5)==true){
      ws.getCell("R"+countxM5).value = Number(this.sum_spliceloss_tM5)
      ws.getCell("R"+countxM5).numFmt = '0.00';
      }else{
       ws.getCell("R"+countxM5).value = 0.00
      ws.getCell("R"+countxM5).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_tM5 > 0 && isNaN(this.sum_splicelossper_tM5)==false && isFinite(this.sum_splicelossper_tM5)==true){
      ws.getCell("S"+countxM5).value = this.sum_splicelossper_tM5/100
      ws.getCell("S"+countxM5).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countxM5).value = 0/100
      ws.getCell("S"+countxM5).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_tM5 > 0 && isNaN(this.sum_totalcutout_tM5)==false && isFinite(this.sum_totalcutout_tM5)==true){
      ws.getCell("U"+countxM5).value = Number(this.sum_totalcutout_tM5)
      ws.getCell("U"+countxM5).numFmt = '0.00';
      }else{
       ws.getCell("U"+countxM5).value = 0.00
      ws.getCell("U"+countxM5).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_tM5 > 0 && isNaN(this.sum_totalcutoutper_tM5)==false && isFinite(this.sum_totalcutoutper_tM5)==true){
      ws.getCell("V"+countxM5).value = this.sum_totalcutoutper_tM5/100
      ws.getCell("V"+countxM5).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countxM5).value = 0.00/100
      ws.getCell("V"+countxM5).numFmt = '0.00%';
      }

       if(this.sum_rement_tM5 > 0 && isNaN(this.sum_rement_tM5)==false && isFinite(this.sum_rement_tM5)==true){
      ws.getCell("X"+countxM5).value = Number(this.sum_rement_tM5)
      ws.getCell("X"+countxM5).numFmt = '0.00';
       }else{
      ws.getCell("X"+countxM5).value = 0.00
      ws.getCell("X"+countxM5).numFmt = '0.00';  
       }

      if(this.sum_rement_per_tM5 > 0 && isNaN(this.sum_rement_per_tM5)==false && isFinite(this.sum_rement_per_tM5)==true){
      ws.getCell("Y"+countxM5).value =  this.sum_rement_per_tM5/100
      ws.getCell("Y"+countxM5).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countxM5).value =  0.00/100
      ws.getCell("Y"+countxM5).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_tM5 > 0 && isNaN(this.sum_totallenthloss_tM5)==false && isFinite(this.sum_totallenthloss_tM5)==true){
      ws.getCell("AA"+countxM5).value = Number(this.sum_totallenthloss_tM5)
      ws.getCell("AA"+countxM5).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countxM5).value = 0
      ws.getCell("AA"+countxM5).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_tM5 = this.sum_totallenthloss_tM5 / this.sum_ttl_marker_tM5 *100
      if(this.last_sum_totallenthloss_per_tM5 > 0 && isNaN(this.last_sum_totallenthloss_per_tM5)==false && isFinite(this.last_sum_totallenthloss_per_tM5)==true){
      ws.getCell("AB"+countxM5).value = this.last_sum_totallenthloss_per_tM5/100
      ws.getCell("AB"+countxM5).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countxM5).value = 0/100
      ws.getCell("AB"+countxM5).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countxM5).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //M6

      if(this.rowexport_type_M5.length == 0){
       
         var countM6 = countM5

       }else{
          var countM6 = countxM5+1
       }
       
      
      

        var countxM6 = countM6 + 1 +Number(this.rowexport_type_M6.length)

    
      this.sum_ttl_marker_tM6 = 0
      this.sum_plug12_tM6 = 0
      this.sum_per_width_loss_tM6 = 0
      this.sum_end1_2_tM6 = 0
      this.sum_endless_length_yd_tM6= 0
      this.sum_avg_end_tM6 = 0
      this.sum_spliceloss_tM6 = 0
      this.sum_splicelossper_tM6 = 0
      this.sum_totalcutout_tM6 = 0
      this.sum_totalcutoutper_tM6 = 0
      this.sum_rement_tM6 = 0
      this.sum_rement_per_tM6 = 0
      this.sum_totallenthloss_tM6 =0 
      this.sum_totallenthlossper_tM6 = 0
      this.summary_widthloss_tM6 = 0
      this.summary_plug12_tM6 = 0
      this.summary_end1_2_tM6 = 0

      
     
      this.sum_all_plug1_2_end_total_tM6 = 0

          if(this.rowexport_type_M6.length < 1){
      count_row_grand++
      
      
      }else{
      
          for (var ax = 0; ax<this.rowexport_type_M6.length; ax++){ //length +M60
      ws.getCell("A"+countM6).value = this.rowexport_type_M6[ax].mu_date

      ws.getCell("B"+countM6).value = this.rowexport_type_M6[ax].gm_table

      ws.getCell("C"+countM6).value = Number(this.rowexport_type_M6[ax].gm_no_short)
      ws.getCell("C"+countM6).numFmt = '0.00';

      ws.getCell("D"+countM6).value = this.rowexport_type_M6[ax].customer

      ws.getCell("E"+countM6).value = this.rowexport_type_M6[ax].gm_no

      ws.getCell("F"+countM6).value = this.rowexport_type_M6[ax].fabric_type

      ws.getCell("G"+countM6).value = Number(this.rowexport_type_M6[ax].obs_width)
      ws.getCell("G"+countM6).numFmt = '0.00';

      ws.getCell("H"+countM6).value = Number(this.rowexport_type_M6[ax].width_inch)
      ws.getCell("H"+countM6).numFmt = '0.00';


      ws.getCell("I"+countM6).value = Number(this.rowexport_type_M6[ax].length_ydb)
      ws.getCell("I"+countM6).numFmt = '0.00';

      ws.getCell("J"+countM6).value = Number(this.rowexport_type_M6[ax].ttl_marker)
      ws.getCell("J"+countM6).numFmt = '0.00';

      ws.getCell("K"+countM6).value = Number(this.rowexport_type_M6[ax].plug12)
      ws.getCell("K"+countM6).numFmt = '0.00';

      ws.getCell("L"+countM6).value = this.rowexport_type_M6[ax].widthloss/100
      ws.getCell("L"+countM6).numFmt = '0.00%';

      ws.getCell("M"+countM6).value = this.rowexport_type_M6[ax].widthloss_code


      ws.getCell("O"+countM6).value = Number(this.rowexport_type_M6[ax].endless_length_yd)
      ws.getCell("O"+countM6).numFmt = '0.00';

      ws.getCell("P"+countM6).value = this.rowexport_type_M6[ax].avg_end/100
      ws.getCell("P"+countM6).numFmt = '0.00%';

      ws.getCell("Q"+countM6).value = this.rowexport_type_M6[ax].avg_end_code
      

      ws.getCell("R"+countM6).value = Number(this.rowexport_type_M6[ax].spliceloss)
      ws.getCell("R"+countM6).numFmt = '0.00';

      ws.getCell("S"+countM6).value = this.rowexport_type_M6[ax].splicelossper/100
      ws.getCell("S"+countM6).numFmt = '0.00%';

      ws.getCell("T"+countM6).value = this.rowexport_type_M6[ax].per_splice_loss_code
   

      ws.getCell("U"+countM6).value = Number(this.rowexport_type_M6[ax].totalcutout)
      ws.getCell("U"+countM6).numFmt = '0.00';

      ws.getCell("V"+countM6).value = this.rowexport_type_M6[ax].totalcutoutper/100
      ws.getCell("V"+countM6).numFmt = '0.00%';

      ws.getCell("W"+countM6).value = this.rowexport_type_M6[ax].total_cut_out_per_code

      ws.getCell("X"+countM6).value = Number(this.rowexport_type_M6[ax].rement)
      ws.getCell("X"+countM6).numFmt = '0.00';

      ws.getCell("Y"+countM6).value = this.rowexport_type_M6[ax].rement_per/100
      ws.getCell("Y"+countM6).numFmt = '0.00%';

      ws.getCell("Z"+countM6).value = this.rowexport_type_M6[ax].remnants_loss_code

      ws.getCell("AA"+countM6).value = Number(this.rowexport_type_M6[ax].totallenthloss)
      ws.getCell("AA"+countM6).numFmt = '0.00';

      ws.getCell("AB"+countM6).value = this.rowexport_type_M6[ax].totallenthlossper/100
      ws.getCell("AB"+countM6).numFmt = '0.00%';
      
      this.sum_plug12_tM6 = 0
      this.sum_ttl_marker_tM6 = Number(this.sum_ttl_marker_tM6) + Number(this.rowexport_type_M6[ax].ttl_marker)
      this.sum_plug12_tM6 = this.rowexport_type_M6[ax].plug12 * this.rowexport_type_M6[ax].ttl_marker
      this.sum_per_width_loss_tM6 = this.rowexport_type_M6[ax].widthloss * this.rowexport_type_M6[ax].ttl_marker
      this.sum_end1_2_tM6 = this.rowexport_type_M6[ax].end1_2 * this.rowexport_type_M6[ax].ttl_marker
      this.sum_endless_length_yd_tM6 = Number(this.sum_endless_length_yd_tM6) + Number(this.rowexport_type_M6[ax].endless_length_yd)
      this.sum_spliceloss_tM6 = Number(this.sum_spliceloss_tM6) + Number(this.rowexport_type_M6[ax].spliceloss)
      this.sum_totalcutout_tM6 = Number(this.sum_totalcutout_tM6) + Number(this.rowexport_type_M6[ax].totalcutout)
      this.sum_rement_tM6 = Number(this.sum_rement_tM6) + Number(this.rowexport_type_M6[ax].rement)
      this.summary_plug12_tM6 = Number(this.summary_plug12_tM6) + Number(this.sum_plug12_tM6)
      this.summary_widthloss_tM6 = Number(this.summary_widthloss_tM6) + Number(this.sum_per_width_loss_tM6)
      this.summary_end1_2_tM6 = Number(this.summary_end1_2_tM6) + Number(this.sum_end1_2_tM6)
      this.sum_totallenthloss_tM6_first = 0
      this.sum_totallenthloss_tM6_first = Number(this.rowexport_type_M6[ax].endless_length_yd) + Number(this.rowexport_type_M6[ax].spliceloss)
      +Number(this.rowexport_type_M6[ax].totalcutout) + Number(this.rowexport_type_M6[ax].rement)
      this.sum_totallenthloss_tM6 = Number(this.sum_totallenthloss_tM6) + Number(this.sum_totallenthloss_tM6_first)

      this.sum_totallenthloss_per_tM6_first = 0
      this.sum_totallenthloss_per_tM6_first = Number(this.rowexport_type_M6[ax].avg_end) + Number(this.rowexport_type_M6[ax].splicelossper)
      +Number(this.rowexport_type_M6[ax].totalcutoutper) + Number(this.rowexport_type_M6[ax].rement_per)
      this.sum_totallenthloss_per_tM6 = Number(this.sum_totallenthloss_tM6) + Number(this.sum_totallenthloss_per_tM6_first)
     this.sum_number_1_end_tM6 = 0 
      this.sum_number_1_end_tM6 = this.rowexport_type_M6[ax].endless_length_yd * 36
      this.sum_number_2_end_tM6 = 0
      this.sum_number_2_end_tM6 = this.rowexport_type_M6[ax].ttl_marker/this.rowexport_type_M6[ax].length_ydb
      this.sum_plug1_2_end_tM6 = 0
      this.sum_plug1_2_end_tM6 = this.sum_number_1_end_tM6/this.sum_number_2_end_tM6

      this.sum_plug1_2_end_total_tM6 = 0
     
      this.sum_plug1_2_end_total_tM6 = this.sum_plug1_2_end_tM6 * this.rowexport_type_M6[ax].ttl_marker
      this.sum_all_plug1_2_end_total_tM6 = Number(this.sum_all_plug1_2_end_total_tM6) + Number(this.sum_plug1_2_end_total_tM6)
      
      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_tM6,
        value_table:"M6"
      })
      ws.getCell("N"+countM6).value = Number(this.sum_plug1_2_end_tM6)
      ws.getCell("N"+countM6).numFmt = '0.00';
      countM6++
         }
         if(this.sum_ttl_marker_tM6 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_tM6
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_tM6 = this.summary_plug12_tM6 / this.sum_ttl_marker_tM6
        if(this.sum_plug12_tM6 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_tM6
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_tM6 =  this.summary_widthloss_tM6 / this.sum_ttl_marker_tM6
        if(this.sum_per_width_loss_tM6 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_tM6
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_tM6 =  this.summary_end1_2_tM6/ this.sum_ttl_marker_tM6 //M
      if(this.sum_end1_2_tM6 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_tM6
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_tM6 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_tM6
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_tM6 = this.sum_endless_length_yd_tM6 / this.sum_ttl_marker_tM6 *100 //P
      if(this.sum_avg_end_tM6 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_tM6
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_tM6 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_tM6
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_tM6 = this.sum_spliceloss_tM6 / this.sum_ttl_marker_tM6 *100
      if(this.sum_splicelossper_tM6 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_tM6
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_tM6 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_tM6
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_tM6 = this.sum_totalcutout_tM6 / this.sum_ttl_marker_tM6 *100
      if(this.sum_totalcutoutper_tM6 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_tM6
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_tM6 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_tM6
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_tM6 = this.sum_rement_tM6 / this.sum_ttl_marker_tM6 *100
  if(this.sum_rement_per_tM6 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_tM6
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_tM6
    if(this.sum_totallenthloss_tM6 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_tM6
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspertM6 = this.sum_totallenthlosstM6 / this.sum_ttl_marker_tM6
    if(this.sum_totallenthlosspertM6 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspertM6
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }


     ws.getCell("A"+countxM6).value = "Total Table M6"  
      ws.mergeCells("A"+countxM6+":"+"I"+countxM6)
      
      if(this.sum_ttl_marker_tM6 > 0 && isNaN(this.sum_ttl_marker_tM6)==false && isFinite(this.sum_ttl_marker_tM6)==true){
      ws.getCell("J"+countxM6).value = Number(this.sum_ttl_marker_tM6)
      ws.getCell("J"+countxM6).numFmt = '0.00';
      }else{
      ws.getCell("J"+countxM6).value = 0
      ws.getCell("J"+countxM6).numFmt = '0.00';
      }

      if(this.sum_plug12_tM6 > 0 && isNaN(this.sum_plug12_tM6)==false && isFinite(this.sum_plug12_tM6)==true){
      ws.getCell("K"+countxM6).value = Number(this.sum_plug12_tM6)
      ws.getCell("K"+countxM6).numFmt = '0.00';
      }else{
      ws.getCell("K"+countxM6).value = 0
      ws.getCell("K"+countxM6).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_tM6 > 0 && isNaN(this.sum_per_width_loss_tM6)==false && isFinite(this.sum_per_width_loss_tM6)==true){
      ws.getCell("L"+countxM6).value = this.sum_per_width_loss_tM6/100
      ws.getCell("L"+countxM6).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countxM6).value = 0.00/100
      ws.getCell("L"+countxM6).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_tM6 = 0
      this.total_result_plug1_2_tM6 = this.sum_all_plug1_2_end_total_tM6 / this.sum_ttl_marker_tM6

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_tM6
      })
      if(this.total_result_plug1_2_tM6 > 0 && isNaN(this.total_result_plug1_2_tM6)==false && isFinite(this.total_result_plug1_2_tM6)==true){
      ws.getCell("N"+countxM6).value = Number(this.total_result_plug1_2_tM6)
      ws.getCell("N"+countxM6).numFmt = '0.00';
      }else{
      ws.getCell("N"+countxM6).value = 0.00
      ws.getCell("N"+countxM6).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_tM6 > 0 && isNaN(this.sum_endless_length_yd_tM6)==false && isFinite(this.sum_endless_length_yd_tM6)==true){
      ws.getCell("O"+countxM6).value = Number(this.sum_endless_length_yd_tM6)
      ws.getCell("O"+countxM6).numFmt = '0.00';
      }else{
      ws.getCell("O"+countxM6).value = 0.00
      ws.getCell("O"+countxM6).numFmt = '0.00';
      }

      if(this.sum_avg_end_tM6 > 0 && isNaN(this.sum_avg_end_tM6)==false && isFinite(this.sum_avg_end_tM6)==true){
      ws.getCell("P"+countxM6).value = this.sum_avg_end_tM6/100
      ws.getCell("P"+countxM6).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countxM6).value = 0/100
      ws.getCell("P"+countxM6).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_tM6 > 0 && isNaN(this.sum_spliceloss_tM6)==false && isFinite(this.sum_spliceloss_tM6)==true){
      ws.getCell("R"+countxM6).value = Number(this.sum_spliceloss_tM6)
      ws.getCell("R"+countxM6).numFmt = '0.00';
      }else{
       ws.getCell("R"+countxM6).value = 0.00
      ws.getCell("R"+countxM6).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_tM6 > 0 && isNaN(this.sum_splicelossper_tM6)==false && isFinite(this.sum_splicelossper_tM6)==true){
      ws.getCell("S"+countxM6).value = this.sum_splicelossper_tM6/100
      ws.getCell("S"+countxM6).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countxM6).value = 0/100
      ws.getCell("S"+countxM6).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_tM6 > 0 && isNaN(this.sum_totalcutout_tM6)==false && isFinite(this.sum_totalcutout_tM6)==true){
      ws.getCell("U"+countxM6).value = Number(this.sum_totalcutout_tM6)
      ws.getCell("U"+countxM6).numFmt = '0.00';
      }else{
       ws.getCell("U"+countxM6).value = 0.00
      ws.getCell("U"+countxM6).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_tM6 > 0 && isNaN(this.sum_totalcutoutper_tM6)==false && isFinite(this.sum_totalcutoutper_tM6)==true){
      ws.getCell("V"+countxM6).value = this.sum_totalcutoutper_tM6/100
      ws.getCell("V"+countxM6).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countxM6).value = 0.00/100
      ws.getCell("V"+countxM6).numFmt = '0.00%';
      }

       if(this.sum_rement_tM6 > 0 && isNaN(this.sum_rement_tM6)==false && isFinite(this.sum_rement_tM6)==true){
      ws.getCell("X"+countxM6).value = Number(this.sum_rement_tM6)
      ws.getCell("X"+countxM6).numFmt = '0.00';
       }else{
      ws.getCell("X"+countxM6).value = 0.00
      ws.getCell("X"+countxM6).numFmt = '0.00';  
       }

      if(this.sum_rement_per_tM6 > 0 && isNaN(this.sum_rement_per_tM6)==false && isFinite(this.sum_rement_per_tM6)==true){
      ws.getCell("Y"+countxM6).value =  this.sum_rement_per_tM6/100
      ws.getCell("Y"+countxM6).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countxM6).value =  0.00/100
      ws.getCell("Y"+countxM6).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_tM6 > 0 && isNaN(this.sum_totallenthloss_tM6)==false && isFinite(this.sum_totallenthloss_tM6)==true){
      ws.getCell("AA"+countxM6).value = Number(this.sum_totallenthloss_tM6)
      ws.getCell("AA"+countxM6).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countxM6).value = 0
      ws.getCell("AA"+countxM6).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_tM6 = this.sum_totallenthloss_tM6 / this.sum_ttl_marker_tM6 *100
      if(this.last_sum_totallenthloss_per_tM6 > 0 && isNaN(this.last_sum_totallenthloss_per_tM6)==false && isFinite(this.last_sum_totallenthloss_per_tM6)==true){
      ws.getCell("AB"+countxM6).value = this.last_sum_totallenthloss_per_tM6/100
      ws.getCell("AB"+countxM6).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countxM6).value = 0/100
      ws.getCell("AB"+countxM6).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countxM6).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //M07

      if(this.rowexport_type_M6.length == 0){
       
         var countM7 = countM6

       }else{
         var countM7 = countxM6+1
       }
       
      

        var countxM7 = countM7 + 1 +Number(this.rowexport_type_M7.length)
        
      
      this.sum_ttl_marker_tM7 = 0
      this.sum_plug12_tM7 = 0
      this.sum_per_width_loss_tM7 = 0
      this.sum_end1_2_tM7 = 0
      this.sum_endless_length_yd_tM7= 0
      this.sum_avg_end_tM7 = 0
      this.sum_spliceloss_tM7 = 0
      this.sum_splicelossper_tM7 = 0
      this.sum_totalcutout_tM7 = 0
      this.sum_totalcutoutper_tM7 = 0
      this.sum_rement_tM7 = 0
      this.sum_rement_per_tM7 = 0
      this.sum_totallenthloss_tM7 =0 
      this.sum_totallenthlossper_tM7 = 0
      this.summary_widthloss_tM7 = 0
      this.summary_plug12_tM7 = 0
      this.summary_end1_2_tM7 = 0

      
     if(this.rowexport_type_M7.length < 1){
      
      count_row_grand++
      
      }else{
      this.sum_all_plug1_2_end_total_tM7 = 0
          for (var ax = 0; ax<this.rowexport_type_M7.length; ax++){ 
      ws.getCell("A"+countM7).value = this.rowexport_type_M7[ax].mu_date

      ws.getCell("B"+countM7).value = this.rowexport_type_M7[ax].gm_table

      ws.getCell("C"+countM7).value = Number(this.rowexport_type_M7[ax].gm_no_short)
      ws.getCell("C"+countM7).numFmt = '0.00';

      ws.getCell("D"+countM7).value = this.rowexport_type_M7[ax].customer

      ws.getCell("E"+countM7).value = this.rowexport_type_M7[ax].gm_no

      ws.getCell("F"+countM7).value = this.rowexport_type_M7[ax].fabric_type

      ws.getCell("G"+countM7).value = Number(this.rowexport_type_M7[ax].obs_width)
      ws.getCell("G"+countM7).numFmt = '0.00';

      ws.getCell("H"+countM7).value = Number(this.rowexport_type_M7[ax].width_inch)
      ws.getCell("H"+countM7).numFmt = '0.00';


      ws.getCell("I"+countM7).value = Number(this.rowexport_type_M7[ax].length_ydb)
      ws.getCell("I"+countM7).numFmt = '0.00';

      ws.getCell("J"+countM7).value = Number(this.rowexport_type_M7[ax].ttl_marker)
      ws.getCell("J"+countM7).numFmt = '0.00';

      ws.getCell("K"+countM7).value = Number(this.rowexport_type_M7[ax].plug12)
      ws.getCell("K"+countM7).numFmt = '0.00';

      ws.getCell("L"+countM7).value = this.rowexport_type_M7[ax].widthloss/100
      ws.getCell("L"+countM7).numFmt = '0.00%';

      ws.getCell("M"+countM7).value = this.rowexport_type_M7[ax].widthloss_code


      ws.getCell("O"+countM7).value = Number(this.rowexport_type_M7[ax].endless_length_yd)
      ws.getCell("O"+countM7).numFmt = '0.00';

      ws.getCell("P"+countM7).value = this.rowexport_type_M7[ax].avg_end/100
      ws.getCell("P"+countM7).numFmt = '0.00%';

      ws.getCell("Q"+countM7).value = this.rowexport_type_M7[ax].avg_end_code
      

      ws.getCell("R"+countM7).value = Number(this.rowexport_type_M7[ax].spliceloss)
      ws.getCell("R"+countM7).numFmt = '0.00';

      ws.getCell("S"+countM7).value = this.rowexport_type_M7[ax].splicelossper/100
      ws.getCell("S"+countM7).numFmt = '0.00%';

      ws.getCell("T"+countM7).value = this.rowexport_type_M7[ax].per_splice_loss_code
   

      ws.getCell("U"+countM7).value = Number(this.rowexport_type_M7[ax].totalcutout)
      ws.getCell("U"+countM7).numFmt = '0.00';

      ws.getCell("V"+countM7).value = this.rowexport_type_M7[ax].totalcutoutper/100
      ws.getCell("V"+countM7).numFmt = '0.00%';

      ws.getCell("W"+countM7).value = this.rowexport_type_M7[ax].total_cut_out_per_code

      ws.getCell("X"+countM7).value = Number(this.rowexport_type_M7[ax].rement)
      ws.getCell("X"+countM7).numFmt = '0.00';

      ws.getCell("Y"+countM7).value = this.rowexport_type_M7[ax].rement_per/100
      ws.getCell("Y"+countM7).numFmt = '0.00%';

      ws.getCell("Z"+countM7).value = this.rowexport_type_M7[ax].remnants_loss_code

      ws.getCell("AA"+countM7).value = Number(this.rowexport_type_M7[ax].totallenthloss)
      ws.getCell("AA"+countM7).numFmt = '0.00';

      ws.getCell("AB"+countM7).value = this.rowexport_type_M7[ax].totallenthlossper/100
      ws.getCell("AB"+countM7).numFmt = '0.00%';
      
      this.sum_plug12_tM7 = 0
      this.sum_ttl_marker_tM7 = Number(this.sum_ttl_marker_tM7) + Number(this.rowexport_type_M7[ax].ttl_marker)
      this.sum_plug12_tM7 = this.rowexport_type_M7[ax].plug12 * this.rowexport_type_M7[ax].ttl_marker
      this.sum_per_width_loss_tM7 = this.rowexport_type_M7[ax].widthloss * this.rowexport_type_M7[ax].ttl_marker
      this.sum_end1_2_tM7 = this.rowexport_type_M7[ax].end1_2 * this.rowexport_type_M7[ax].ttl_marker
      this.sum_endless_length_yd_tM7 = Number(this.sum_endless_length_yd_tM7) + Number(this.rowexport_type_M7[ax].endless_length_yd)
      this.sum_spliceloss_tM7 = Number(this.sum_spliceloss_tM7) + Number(this.rowexport_type_M7[ax].spliceloss)
      this.sum_totalcutout_tM7 = Number(this.sum_totalcutout_tM7) + Number(this.rowexport_type_M7[ax].totalcutout)
      this.sum_rement_tM7 = Number(this.sum_rement_tM7) + Number(this.rowexport_type_M7[ax].rement)
      this.summary_plug12_tM7 = Number(this.summary_plug12_tM7) + Number(this.sum_plug12_tM7)
      this.summary_widthloss_tM7 = Number(this.summary_widthloss_tM7) + Number(this.sum_per_width_loss_tM7)
      this.summary_end1_2_tM7 = Number(this.summary_end1_2_tM7) + Number(this.sum_end1_2_tM7)
      this.sum_totallenthloss_tM7_first = 0
      this.sum_totallenthloss_tM7_first = Number(this.rowexport_type_M7[ax].endless_length_yd) + Number(this.rowexport_type_M7[ax].spliceloss)
      +Number(this.rowexport_type_M7[ax].totalcutout) + Number(this.rowexport_type_M7[ax].rement)
      this.sum_totallenthloss_tM7 = Number(this.sum_totallenthloss_tM7) + Number(this.sum_totallenthloss_tM7_first)

      this.sum_totallenthloss_per_tM7_first = 0
      this.sum_totallenthloss_per_tM7_first = Number(this.rowexport_type_M7[ax].avg_end) + Number(this.rowexport_type_M7[ax].splicelossper)
      +Number(this.rowexport_type_M7[ax].totalcutoutper) + Number(this.rowexport_type_M7[ax].rement_per)
      this.sum_totallenthloss_per_tM7 = Number(this.sum_totallenthloss_tM7) + Number(this.sum_totallenthloss_per_tM7_first)
     this.sum_number_1_end_tM7 = 0 
      this.sum_number_1_end_tM7 = this.rowexport_type_M7[ax].endless_length_yd * 36
      this.sum_number_2_end_tM7 = 0
      this.sum_number_2_end_tM7 = this.rowexport_type_M7[ax].ttl_marker/this.rowexport_type_M7[ax].length_ydb
      this.sum_plug1_2_end_tM7 = 0
      this.sum_plug1_2_end_tM7 = this.sum_number_1_end_tM7/this.sum_number_2_end_tM7

      this.sum_plug1_2_end_total_tM7 = 0
     
      this.sum_plug1_2_end_total_tM7 = this.sum_plug1_2_end_tM7 * this.rowexport_type_M7[ax].ttl_marker
      this.sum_all_plug1_2_end_total_tM7 = Number(this.sum_all_plug1_2_end_total_tM7) + Number(this.sum_plug1_2_end_total_tM7)
      
      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_tM7,
        value_table:"M7"
      })
      ws.getCell("N"+countM7).value = Number(this.sum_plug1_2_end_tM7)
      ws.getCell("N"+countM7).numFmt = '0.00';
      countM7++
         }
         if(this.sum_ttl_marker_tM7 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_tM7
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_tM7 = this.summary_plug12_tM7 / this.sum_ttl_marker_tM7
        if(this.sum_plug12_tM7 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_tM7
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_tM7 =  this.summary_widthloss_tM7 / this.sum_ttl_marker_tM7
        if(this.sum_per_width_loss_tM7 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_tM7
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_tM7 =  this.summary_end1_2_tM7/ this.sum_ttl_marker_tM7 //M
      if(this.sum_end1_2_tM7 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_tM7
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_tM7 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_tM7
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_tM7 = this.sum_endless_length_yd_tM7 / this.sum_ttl_marker_tM7 *100 //P
      if(this.sum_avg_end_tM7 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_tM7
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_tM7 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_tM7
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_tM7 = this.sum_spliceloss_tM7 / this.sum_ttl_marker_tM7 *100
      if(this.sum_splicelossper_tM7 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_tM7
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_tM7 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_tM7
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_tM7 = this.sum_totalcutout_tM7 / this.sum_ttl_marker_tM7 *100
      if(this.sum_totalcutoutper_tM7 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_tM7
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_tM7 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_tM7
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_tM7 = this.sum_rement_tM7 / this.sum_ttl_marker_tM7 *100
  if(this.sum_rement_per_tM7 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_tM7
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_tM7
    if(this.sum_totallenthloss_tM7 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_tM7
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspertM7 = this.sum_totallenthlosstM7 / this.sum_ttl_marker_tM7
    if(this.sum_totallenthlosspertM7 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspertM7
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }


     ws.getCell("A"+countxM7).value = "Total Table M7"  
      ws.mergeCells("A"+countxM7+":"+"I"+countxM7)
      
      if(this.sum_ttl_marker_tM7 > 0 && isNaN(this.sum_ttl_marker_tM7)==false && isFinite(this.sum_ttl_marker_tM7)==true){
      ws.getCell("J"+countxM7).value = Number(this.sum_ttl_marker_tM7)
      ws.getCell("J"+countxM7).numFmt = '0.00';
      }else{
      ws.getCell("J"+countxM7).value = 0
      ws.getCell("J"+countxM7).numFmt = '0.00';
      }

      if(this.sum_plug12_tM7 > 0 && isNaN(this.sum_plug12_tM7)==false && isFinite(this.sum_plug12_tM7)==true){
      ws.getCell("K"+countxM7).value = Number(this.sum_plug12_tM7)
      ws.getCell("K"+countxM7).numFmt = '0.00';
      }else{
      ws.getCell("K"+countxM7).value = 0
      ws.getCell("K"+countxM7).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_tM7 > 0 && isNaN(this.sum_per_width_loss_tM7)==false && isFinite(this.sum_per_width_loss_tM7)==true){
      ws.getCell("L"+countxM7).value = this.sum_per_width_loss_tM7/100
      ws.getCell("L"+countxM7).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countxM7).value = 0.00/100
      ws.getCell("L"+countxM7).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_tM7 = 0
      this.total_result_plug1_2_tM7 = this.sum_all_plug1_2_end_total_tM7 / this.sum_ttl_marker_tM7

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_tM7
      })
      if(this.total_result_plug1_2_tM7 > 0 && isNaN(this.total_result_plug1_2_tM7)==false && isFinite(this.total_result_plug1_2_tM7)==true){
      ws.getCell("N"+countxM7).value = Number(this.total_result_plug1_2_tM7)
      ws.getCell("N"+countxM7).numFmt = '0.00';
      }else{
      ws.getCell("N"+countxM7).value = 0.00
      ws.getCell("N"+countxM7).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_tM7 > 0 && isNaN(this.sum_endless_length_yd_tM7)==false && isFinite(this.sum_endless_length_yd_tM7)==true){
      ws.getCell("O"+countxM7).value = Number(this.sum_endless_length_yd_tM7)
      ws.getCell("O"+countxM7).numFmt = '0.00';
      }else{
      ws.getCell("O"+countxM7).value = 0.00
      ws.getCell("O"+countxM7).numFmt = '0.00';
      }

      if(this.sum_avg_end_tM7 > 0 && isNaN(this.sum_avg_end_tM7)==false && isFinite(this.sum_avg_end_tM7)==true){
      ws.getCell("P"+countxM7).value = this.sum_avg_end_tM7/100
      ws.getCell("P"+countxM7).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countxM7).value = 0/100
      ws.getCell("P"+countxM7).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_tM7 > 0 && isNaN(this.sum_spliceloss_tM7)==false && isFinite(this.sum_spliceloss_tM7)==true){
      ws.getCell("R"+countxM7).value = Number(this.sum_spliceloss_tM7)
      ws.getCell("R"+countxM7).numFmt = '0.00';
      }else{
       ws.getCell("R"+countxM7).value = 0.00
      ws.getCell("R"+countxM7).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_tM7 > 0 && isNaN(this.sum_splicelossper_tM7)==false && isFinite(this.sum_splicelossper_tM7)==true){
      ws.getCell("S"+countxM7).value = this.sum_splicelossper_tM7/100
      ws.getCell("S"+countxM7).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countxM7).value = 0/100
      ws.getCell("S"+countxM7).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_tM7 > 0 && isNaN(this.sum_totalcutout_tM7)==false && isFinite(this.sum_totalcutout_tM7)==true){
      ws.getCell("U"+countxM7).value = Number(this.sum_totalcutout_tM7)
      ws.getCell("U"+countxM7).numFmt = '0.00';
      }else{
       ws.getCell("U"+countxM7).value = 0.00
      ws.getCell("U"+countxM7).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_tM7 > 0 && isNaN(this.sum_totalcutoutper_tM7)==false && isFinite(this.sum_totalcutoutper_tM7)==true){
      ws.getCell("V"+countxM7).value = this.sum_totalcutoutper_tM7/100
      ws.getCell("V"+countxM7).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countxM7).value = 0.00/100
      ws.getCell("V"+countxM7).numFmt = '0.00%';
      }

       if(this.sum_rement_tM7 > 0 && isNaN(this.sum_rement_tM7)==false && isFinite(this.sum_rement_tM7)==true){
      ws.getCell("X"+countxM7).value = Number(this.sum_rement_tM7)
      ws.getCell("X"+countxM7).numFmt = '0.00';
       }else{
      ws.getCell("X"+countxM7).value = 0.00
      ws.getCell("X"+countxM7).numFmt = '0.00';  
       }

      if(this.sum_rement_per_tM7 > 0 && isNaN(this.sum_rement_per_tM7)==false && isFinite(this.sum_rement_per_tM7)==true){
      ws.getCell("Y"+countxM7).value =  this.sum_rement_per_tM7/100
      ws.getCell("Y"+countxM7).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countxM7).value =  0.00/100
      ws.getCell("Y"+countxM7).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_tM7 > 0 && isNaN(this.sum_totallenthloss_tM7)==false && isFinite(this.sum_totallenthloss_tM7)==true){
      ws.getCell("AA"+countxM7).value = Number(this.sum_totallenthloss_tM7)
      ws.getCell("AA"+countxM7).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countxM7).value = 0
      ws.getCell("AA"+countxM7).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_tM7 = this.sum_totallenthloss_tM7 / this.sum_ttl_marker_tM7 *100
      if(this.last_sum_totallenthloss_per_tM7 > 0 && isNaN(this.last_sum_totallenthloss_per_tM7)==false && isFinite(this.last_sum_totallenthloss_per_tM7)==true){
      ws.getCell("AB"+countxM7).value = this.last_sum_totallenthloss_per_tM7/100
      ws.getCell("AB"+countxM7).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countxM7).value = 0/100
      ws.getCell("AB"+countxM7).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countxM7).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //M08

       if(this.rowexport_type_M7.length == 0){
       
         var countM8 = countM7

       }else{
         var countM8 = countxM7+1
       }

        var countxM8 = countM8 + 1 +Number(this.rowexport_type_M8.length)
      
      this.sum_ttl_marker_tM8 = 0
      this.sum_plug12_tM8 = 0
      this.sum_per_width_loss_tM8 = 0
      this.sum_end1_2_tM8 = 0
      this.sum_endless_length_yd_tM8= 0
      this.sum_avg_end_tM8 = 0
      this.sum_spliceloss_tM8 = 0
      this.sum_splicelossper_tM8 = 0
      this.sum_totalcutout_tM8 = 0
      this.sum_totalcutoutper_tM8 = 0
      this.sum_rement_tM8 = 0
      this.sum_rement_per_tM8 = 0
      this.sum_totallenthloss_tM8 =0 
      this.sum_totallenthlossper_tM8 = 0
      this.summary_widthloss_tM8 = 0
      this.summary_plug12_tM8 = 0
      this.summary_end1_2_tM8 = 0

      
     
      this.sum_all_plug1_2_end_total_tM8 = 0

      if(this.rowexport_type_M8.length < 1){
      
      count_row_grand++
      
      }else{
          for (var ax = 0; ax<this.rowexport_type_M8.length; ax++){ 
      ws.getCell("A"+countM8).value = this.rowexport_type_M8[ax].mu_date

      ws.getCell("B"+countM8).value = this.rowexport_type_M8[ax].gm_table

      ws.getCell("C"+countM8).value = Number(this.rowexport_type_M8[ax].gm_no_short)
      ws.getCell("C"+countM8).numFmt = '0.00';

      ws.getCell("D"+countM8).value = this.rowexport_type_M8[ax].customer

      ws.getCell("E"+countM8).value = this.rowexport_type_M8[ax].gm_no

      ws.getCell("F"+countM8).value = this.rowexport_type_M8[ax].fabric_type

      ws.getCell("G"+countM8).value = Number(this.rowexport_type_M8[ax].obs_width)
      ws.getCell("G"+countM8).numFmt = '0.00';

      ws.getCell("H"+countM8).value = Number(this.rowexport_type_M8[ax].width_inch)
      ws.getCell("H"+countM8).numFmt = '0.00';


      ws.getCell("I"+countM8).value = Number(this.rowexport_type_M8[ax].length_ydb)
      ws.getCell("I"+countM8).numFmt = '0.00';

      ws.getCell("J"+countM8).value = Number(this.rowexport_type_M8[ax].ttl_marker)
      ws.getCell("J"+countM8).numFmt = '0.00';

      ws.getCell("K"+countM8).value = Number(this.rowexport_type_M8[ax].plug12)
      ws.getCell("K"+countM8).numFmt = '0.00';

      ws.getCell("L"+countM8).value = this.rowexport_type_M8[ax].widthloss/100
      ws.getCell("L"+countM8).numFmt = '0.00%';

      ws.getCell("M"+countM8).value = this.rowexport_type_M8[ax].widthloss_code


      ws.getCell("O"+countM8).value = Number(this.rowexport_type_M8[ax].endless_length_yd)
      ws.getCell("O"+countM8).numFmt = '0.00';

      ws.getCell("P"+countM8).value = this.rowexport_type_M8[ax].avg_end/100
      ws.getCell("P"+countM8).numFmt = '0.00%';

      ws.getCell("Q"+countM8).value = this.rowexport_type_M8[ax].avg_end_code
      

      ws.getCell("R"+countM8).value = Number(this.rowexport_type_M8[ax].spliceloss)
      ws.getCell("R"+countM8).numFmt = '0.00';

      ws.getCell("S"+countM8).value = this.rowexport_type_M8[ax].splicelossper/100
      ws.getCell("S"+countM8).numFmt = '0.00%';

      ws.getCell("T"+countM8).value = this.rowexport_type_M8[ax].per_splice_loss_code
   

      ws.getCell("U"+countM8).value = Number(this.rowexport_type_M8[ax].totalcutout)
      ws.getCell("U"+countM8).numFmt = '0.00';

      ws.getCell("V"+countM8).value = this.rowexport_type_M8[ax].totalcutoutper/100
      ws.getCell("V"+countM8).numFmt = '0.00%';

      ws.getCell("W"+countM8).value = this.rowexport_type_M8[ax].total_cut_out_per_code

      ws.getCell("X"+countM8).value = Number(this.rowexport_type_M8[ax].rement)
      ws.getCell("X"+countM8).numFmt = '0.00';

      ws.getCell("Y"+countM8).value = this.rowexport_type_M8[ax].rement_per/100
      ws.getCell("Y"+countM8).numFmt = '0.00%';

      ws.getCell("Z"+countM8).value = this.rowexport_type_M8[ax].remnants_loss_code

      ws.getCell("AA"+countM8).value = Number(this.rowexport_type_M8[ax].totallenthloss)
      ws.getCell("AA"+countM8).numFmt = '0.00';

      ws.getCell("AB"+countM8).value = this.rowexport_type_M8[ax].totallenthlossper/100
      ws.getCell("AB"+countM8).numFmt = '0.00%';
      
      this.sum_plug12_tM8 = 0
      this.sum_ttl_marker_tM8 = Number(this.sum_ttl_marker_tM8) + Number(this.rowexport_type_M8[ax].ttl_marker)
      this.sum_plug12_tM8 = this.rowexport_type_M8[ax].plug12 * this.rowexport_type_M8[ax].ttl_marker
      this.sum_per_width_loss_tM8 = this.rowexport_type_M8[ax].widthloss * this.rowexport_type_M8[ax].ttl_marker
      this.sum_end1_2_tM8 = this.rowexport_type_M8[ax].end1_2 * this.rowexport_type_M8[ax].ttl_marker
      this.sum_endless_length_yd_tM8 = Number(this.sum_endless_length_yd_tM8) + Number(this.rowexport_type_M8[ax].endless_length_yd)
      this.sum_spliceloss_tM8 = Number(this.sum_spliceloss_tM8) + Number(this.rowexport_type_M8[ax].spliceloss)
      this.sum_totalcutout_tM8 = Number(this.sum_totalcutout_tM8) + Number(this.rowexport_type_M8[ax].totalcutout)
      this.sum_rement_tM8 = Number(this.sum_rement_tM8) + Number(this.rowexport_type_M8[ax].rement)
      this.summary_plug12_tM8 = Number(this.summary_plug12_tM8) + Number(this.sum_plug12_tM8)
      this.summary_widthloss_tM8 = Number(this.summary_widthloss_tM8) + Number(this.sum_per_width_loss_tM8)
      this.summary_end1_2_tM8 = Number(this.summary_end1_2_tM8) + Number(this.sum_end1_2_tM8)
      this.sum_totallenthloss_tM8_first = 0
      this.sum_totallenthloss_tM8_first = Number(this.rowexport_type_M8[ax].endless_length_yd) + Number(this.rowexport_type_M8[ax].spliceloss)
      +Number(this.rowexport_type_M8[ax].totalcutout) + Number(this.rowexport_type_M8[ax].rement)
      this.sum_totallenthloss_tM8 = Number(this.sum_totallenthloss_tM8) + Number(this.sum_totallenthloss_tM8_first)

      this.sum_totallenthloss_per_tM8_first = 0
      this.sum_totallenthloss_per_tM8_first = Number(this.rowexport_type_M8[ax].avg_end) + Number(this.rowexport_type_M8[ax].splicelossper)
      +Number(this.rowexport_type_M8[ax].totalcutoutper) + Number(this.rowexport_type_M8[ax].rement_per)
      this.sum_totallenthloss_per_tM8 = Number(this.sum_totallenthloss_tM8) + Number(this.sum_totallenthloss_per_tM8_first)
     this.sum_number_1_end_tM8 = 0 
      this.sum_number_1_end_tM8 = this.rowexport_type_M8[ax].endless_length_yd * 36
      this.sum_number_2_end_tM8 = 0
      this.sum_number_2_end_tM8 = this.rowexport_type_M8[ax].ttl_marker/this.rowexport_type_M8[ax].length_ydb
      this.sum_plug1_2_end_tM8 = 0
      this.sum_plug1_2_end_tM8 = this.sum_number_1_end_tM8/this.sum_number_2_end_tM8

      this.sum_plug1_2_end_total_tM8 = 0
     
      this.sum_plug1_2_end_total_tM8 = this.sum_plug1_2_end_tM8 * this.rowexport_type_M8[ax].ttl_marker
      this.sum_all_plug1_2_end_total_tM8 = Number(this.sum_all_plug1_2_end_total_tM8) + Number(this.sum_plug1_2_end_total_tM8)
      
      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_tM8,
        value_table:"M8"
      })
      ws.getCell("N"+countM8).value = Number(this.sum_plug1_2_end_tM8)
      ws.getCell("N"+countM8).numFmt = '0.00';
      countM8++
         }
         if(this.sum_ttl_marker_tM8 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_tM8
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_tM8 = this.summary_plug12_tM8 / this.sum_ttl_marker_tM8
        if(this.sum_plug12_tM8 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_tM8
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_tM8 =  this.summary_widthloss_tM8 / this.sum_ttl_marker_tM8
        if(this.sum_per_width_loss_tM8 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_tM8
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_tM8 =  this.summary_end1_2_tM8/ this.sum_ttl_marker_tM8 //M
      if(this.sum_end1_2_tM8 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_tM8
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_tM8 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_tM8
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_tM8 = this.sum_endless_length_yd_tM8 / this.sum_ttl_marker_tM8 *100 //P
      if(this.sum_avg_end_tM8 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_tM8
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_tM8 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_tM8
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_tM8 = this.sum_spliceloss_tM8 / this.sum_ttl_marker_tM8 *100
      if(this.sum_splicelossper_tM8 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_tM8
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_tM8 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_tM8
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_tM8 = this.sum_totalcutout_tM8 / this.sum_ttl_marker_tM8 *100
      if(this.sum_totalcutoutper_tM8 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_tM8
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_tM8 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_tM8
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_tM8 = this.sum_rement_tM8 / this.sum_ttl_marker_tM8 *100
  if(this.sum_rement_per_tM8 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_tM8
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_tM8
    if(this.sum_totallenthloss_tM8 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_tM8
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspertM8 = this.sum_totallenthlosstM8 / this.sum_ttl_marker_tM8
    if(this.sum_totallenthlosspertM8 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspertM8
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }


     ws.getCell("A"+countxM8).value = "Total Table M8"  
      ws.mergeCells("A"+countxM8+":"+"I"+countxM8)
      
      if(this.sum_ttl_marker_tM8 > 0 && isNaN(this.sum_ttl_marker_tM8)==false && isFinite(this.sum_ttl_marker_tM8)==true){
      ws.getCell("J"+countxM8).value = Number(this.sum_ttl_marker_tM8)
      ws.getCell("J"+countxM8).numFmt = '0.00';
      }else{
      ws.getCell("J"+countxM8).value = 0
      ws.getCell("J"+countxM8).numFmt = '0.00';
      }

      if(this.sum_plug12_tM8 > 0 && isNaN(this.sum_plug12_tM8)==false && isFinite(this.sum_plug12_tM8)==true){
      ws.getCell("K"+countxM8).value = Number(this.sum_plug12_tM8)
      ws.getCell("K"+countxM8).numFmt = '0.00';
      }else{
      ws.getCell("K"+countxM8).value = 0
      ws.getCell("K"+countxM8).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_tM8 > 0 && isNaN(this.sum_per_width_loss_tM8)==false && isFinite(this.sum_per_width_loss_tM8)==true){
      ws.getCell("L"+countxM8).value = this.sum_per_width_loss_tM8/100
      ws.getCell("L"+countxM8).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countxM8).value = 0.00/100
      ws.getCell("L"+countxM8).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_tM8 = 0
      this.total_result_plug1_2_tM8 = this.sum_all_plug1_2_end_total_tM8 / this.sum_ttl_marker_tM8

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_tM8
      })
      if(this.total_result_plug1_2_tM8 > 0 && isNaN(this.total_result_plug1_2_tM8)==false && isFinite(this.total_result_plug1_2_tM8)==true){
      ws.getCell("N"+countxM8).value = Number(this.total_result_plug1_2_tM8)
      ws.getCell("N"+countxM8).numFmt = '0.00';
      }else{
      ws.getCell("N"+countxM8).value = 0.00
      ws.getCell("N"+countxM8).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_tM8 > 0 && isNaN(this.sum_endless_length_yd_tM8)==false && isFinite(this.sum_endless_length_yd_tM8)==true){
      ws.getCell("O"+countxM8).value = Number(this.sum_endless_length_yd_tM8)
      ws.getCell("O"+countxM8).numFmt = '0.00';
      }else{
      ws.getCell("O"+countxM8).value = 0.00
      ws.getCell("O"+countxM8).numFmt = '0.00';
      }

      if(this.sum_avg_end_tM8 > 0 && isNaN(this.sum_avg_end_tM8)==false && isFinite(this.sum_avg_end_tM8)==true){
      ws.getCell("P"+countxM8).value = this.sum_avg_end_tM8/100
      ws.getCell("P"+countxM8).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countxM8).value = 0/100
      ws.getCell("P"+countxM8).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_tM8 > 0 && isNaN(this.sum_spliceloss_tM8)==false && isFinite(this.sum_spliceloss_tM8)==true){
      ws.getCell("R"+countxM8).value = Number(this.sum_spliceloss_tM8)
      ws.getCell("R"+countxM8).numFmt = '0.00';
      }else{
       ws.getCell("R"+countxM8).value = 0.00
      ws.getCell("R"+countxM8).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_tM8 > 0 && isNaN(this.sum_splicelossper_tM8)==false && isFinite(this.sum_splicelossper_tM8)==true){
      ws.getCell("S"+countxM8).value = this.sum_splicelossper_tM8/100
      ws.getCell("S"+countxM8).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countxM8).value = 0/100
      ws.getCell("S"+countxM8).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_tM8 > 0 && isNaN(this.sum_totalcutout_tM8)==false && isFinite(this.sum_totalcutout_tM8)==true){
      ws.getCell("U"+countxM8).value = Number(this.sum_totalcutout_tM8)
      ws.getCell("U"+countxM8).numFmt = '0.00';
      }else{
       ws.getCell("U"+countxM8).value = 0.00
      ws.getCell("U"+countxM8).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_tM8 > 0 && isNaN(this.sum_totalcutoutper_tM8)==false && isFinite(this.sum_totalcutoutper_tM8)==true){
      ws.getCell("V"+countxM8).value = this.sum_totalcutoutper_tM8/100
      ws.getCell("V"+countxM8).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countxM8).value = 0.00/100
      ws.getCell("V"+countxM8).numFmt = '0.00%';
      }

       if(this.sum_rement_tM8 > 0 && isNaN(this.sum_rement_tM8)==false && isFinite(this.sum_rement_tM8)==true){
      ws.getCell("X"+countxM8).value = Number(this.sum_rement_tM8)
      ws.getCell("X"+countxM8).numFmt = '0.00';
       }else{
      ws.getCell("X"+countxM8).value = 0.00
      ws.getCell("X"+countxM8).numFmt = '0.00';  
       }

      if(this.sum_rement_per_tM8 > 0 && isNaN(this.sum_rement_per_tM8)==false && isFinite(this.sum_rement_per_tM8)==true){
      ws.getCell("Y"+countxM8).value =  this.sum_rement_per_tM8/100
      ws.getCell("Y"+countxM8).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countxM8).value =  0.00/100
      ws.getCell("Y"+countxM8).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_tM8 > 0 && isNaN(this.sum_totallenthloss_tM8)==false && isFinite(this.sum_totallenthloss_tM8)==true){
      ws.getCell("AA"+countxM8).value = Number(this.sum_totallenthloss_tM8)
      ws.getCell("AA"+countxM8).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countxM8).value = 0
      ws.getCell("AA"+countxM8).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_tM8 = this.sum_totallenthloss_tM8 / this.sum_ttl_marker_tM8 *100
      if(this.last_sum_totallenthloss_per_tM8 > 0 && isNaN(this.last_sum_totallenthloss_per_tM8)==false && isFinite(this.last_sum_totallenthloss_per_tM8)==true){
      ws.getCell("AB"+countxM8).value = this.last_sum_totallenthloss_per_tM8/100
      ws.getCell("AB"+countxM8).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countxM8).value = 0/100
      ws.getCell("AB"+countxM8).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countxM8).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //M09

      if(this.rowexport_type_M8.length == 0){
       
         var countM9 = countM8

       }else{
         var countM9 = countxM8+1
       }
      

        var countxM9 = countM9 + 1 +Number(this.rowexport_type_M9.length)
      
      this.sum_ttl_marker_tM9 = 0
      this.sum_plug12_tM9 = 0
      this.sum_per_width_loss_tM9 = 0
      this.sum_end1_2_tM9 = 0
      this.sum_endless_length_yd_tM9= 0
      this.sum_avg_end_tM9 = 0
      this.sum_spliceloss_tM9 = 0
      this.sum_splicelossper_tM9 = 0
      this.sum_totalcutout_tM9 = 0
      this.sum_totalcutoutper_tM9 = 0
      this.sum_rement_tM9 = 0
      this.sum_rement_per_tM9 = 0
      this.sum_totallenthloss_tM9 =0 
      this.sum_totallenthlossper_tM9 = 0
      this.summary_widthloss_tM9 = 0
      this.summary_plug12_tM9 = 0
      this.summary_end1_2_tM9 = 0

      
     
      this.sum_all_plug1_2_end_total_tM9 = 0

      if(this.rowexport_type_M9.length < 1){
      count_row_grand++
      
      
      }else{
    
          for (var ax = 0; ax<this.rowexport_type_M9.length; ax++){ 
      ws.getCell("A"+countM9).value = this.rowexport_type_M9[ax].mu_date

      ws.getCell("B"+countM9).value = this.rowexport_type_M9[ax].gm_table

      ws.getCell("C"+countM9).value = Number(this.rowexport_type_M9[ax].gm_no_short)
      ws.getCell("C"+countM9).numFmt = '0.00';

      ws.getCell("D"+countM9).value = this.rowexport_type_M9[ax].customer

      ws.getCell("E"+countM9).value = this.rowexport_type_M9[ax].gm_no

      ws.getCell("F"+countM9).value = this.rowexport_type_M9[ax].fabric_type

      ws.getCell("G"+countM9).value = Number(this.rowexport_type_M9[ax].obs_width)
      ws.getCell("G"+countM9).numFmt = '0.00';

      ws.getCell("H"+countM9).value = Number(this.rowexport_type_M9[ax].width_inch)
      ws.getCell("H"+countM9).numFmt = '0.00';


      ws.getCell("I"+countM9).value = Number(this.rowexport_type_M9[ax].length_ydb)
      ws.getCell("I"+countM9).numFmt = '0.00';

      ws.getCell("J"+countM9).value = Number(this.rowexport_type_M9[ax].ttl_marker)
      ws.getCell("J"+countM9).numFmt = '0.00';

      ws.getCell("K"+countM9).value = Number(this.rowexport_type_M9[ax].plug12)
      ws.getCell("K"+countM9).numFmt = '0.00';

      ws.getCell("L"+countM9).value = this.rowexport_type_M9[ax].widthloss/100
      ws.getCell("L"+countM9).numFmt = '0.00%';

      ws.getCell("M"+countM9).value = this.rowexport_type_M9[ax].widthloss_code


      ws.getCell("O"+countM9).value = Number(this.rowexport_type_M9[ax].endless_length_yd)
      ws.getCell("O"+countM9).numFmt = '0.00';

      ws.getCell("P"+countM9).value = this.rowexport_type_M9[ax].avg_end/100
      ws.getCell("P"+countM9).numFmt = '0.00%';

      ws.getCell("Q"+countM9).value = this.rowexport_type_M9[ax].avg_end_code
      
   
      ws.getCell("R"+countM9).value = Number(this.rowexport_type_M9[ax].spliceloss)
      ws.getCell("R"+countM9).numFmt = '0.00';

      ws.getCell("S"+countM9).value = this.rowexport_type_M9[ax].splicelossper/100
      ws.getCell("S"+countM9).numFmt = '0.00%';

      ws.getCell("T"+countM9).value = this.rowexport_type_M9[ax].per_splice_loss_code
   

      ws.getCell("U"+countM9).value = Number(this.rowexport_type_M9[ax].totalcutout)
      ws.getCell("U"+countM9).numFmt = '0.00';

      ws.getCell("V"+countM9).value = this.rowexport_type_M9[ax].totalcutoutper/100
      ws.getCell("V"+countM9).numFmt = '0.00%';

      ws.getCell("W"+countM9).value = this.rowexport_type_M9[ax].total_cut_out_per_code

      ws.getCell("X"+countM9).value = Number(this.rowexport_type_M9[ax].rement)
      ws.getCell("X"+countM9).numFmt = '0.00';

      ws.getCell("Y"+countM9).value = this.rowexport_type_M9[ax].rement_per/100
      ws.getCell("Y"+countM9).numFmt = '0.00%';

      ws.getCell("Z"+countM9).value = this.rowexport_type_M9[ax].remnants_loss_code

      ws.getCell("AA"+countM9).value = Number(this.rowexport_type_M9[ax].totallenthloss)
      ws.getCell("AA"+countM9).numFmt = '0.00';

      ws.getCell("AB"+countM9).value = this.rowexport_type_M9[ax].totallenthlossper/100
      ws.getCell("AB"+countM9).numFmt = '0.00%';
      
      
      
      this.sum_plug12_tM9 = 0
      this.sum_ttl_marker_tM9 = Number(this.sum_ttl_marker_tM9) + Number(this.rowexport_type_M9[ax].ttl_marker)
      this.sum_plug12_tM9 = this.rowexport_type_M9[ax].plug12 * this.rowexport_type_M9[ax].ttl_marker
      this.sum_per_width_loss_tM9 = this.rowexport_type_M9[ax].widthloss * this.rowexport_type_M9[ax].ttl_marker
      this.sum_end1_2_tM9 = this.rowexport_type_M9[ax].end1_2 * this.rowexport_type_M9[ax].ttl_marker
      this.sum_endless_length_yd_tM9 = Number(this.sum_endless_length_yd_tM9) + Number(this.rowexport_type_M9[ax].endless_length_yd)
      this.sum_spliceloss_tM9 = Number(this.sum_spliceloss_tM9) + Number(this.rowexport_type_M9[ax].spliceloss)
      this.sum_totalcutout_tM9 = Number(this.sum_totalcutout_tM9) + Number(this.rowexport_type_M9[ax].totalcutout)
      this.sum_rement_tM9 = Number(this.sum_rement_tM9) + Number(this.rowexport_type_M9[ax].rement)
      this.summary_plug12_tM9 = Number(this.summary_plug12_tM9) + Number(this.sum_plug12_tM9)
      this.summary_widthloss_tM9 = Number(this.summary_widthloss_tM9) + Number(this.sum_per_width_loss_tM9)
      this.summary_end1_2_tM9 = Number(this.summary_end1_2_tM9) + Number(this.sum_end1_2_tM9)
      this.sum_totallenthloss_tM9_first = 0
      this.sum_totallenthloss_tM9_first = Number(this.rowexport_type_M9[ax].endless_length_yd) + Number(this.rowexport_type_M9[ax].spliceloss)
      +Number(this.rowexport_type_M9[ax].totalcutout) + Number(this.rowexport_type_M9[ax].rement)
      this.sum_totallenthloss_tM9 = Number(this.sum_totallenthloss_tM9) + Number(this.sum_totallenthloss_tM9_first)

      this.sum_totallenthloss_per_tM9_first = 0
      this.sum_totallenthloss_per_tM9_first = Number(this.rowexport_type_M9[ax].avg_end) + Number(this.rowexport_type_M9[ax].splicelossper)
      +Number(this.rowexport_type_M9[ax].totalcutoutper) + Number(this.rowexport_type_M9[ax].rement_per)
      this.sum_totallenthloss_per_tM9 = Number(this.sum_totallenthloss_tM9) + Number(this.sum_totallenthloss_per_tM9_first)
     this.sum_number_1_end_tM9 = 0 
      this.sum_number_1_end_tM9 = this.rowexport_type_M9[ax].endless_length_yd * 36
      this.sum_number_2_end_tM9 = 0
      this.sum_number_2_end_tM9 = this.rowexport_type_M9[ax].ttl_marker/this.rowexport_type_M9[ax].length_ydb
      this.sum_plug1_2_end_tM9 = 0
      this.sum_plug1_2_end_tM9 = this.sum_number_1_end_tM9/this.sum_number_2_end_tM9

      this.sum_plug1_2_end_total_tM9 = 0
     
      this.sum_plug1_2_end_total_tM9 = this.sum_plug1_2_end_tM9 * this.rowexport_type_M9[ax].ttl_marker
      this.sum_all_plug1_2_end_total_tM9 = Number(this.sum_all_plug1_2_end_total_tM9) + Number(this.sum_plug1_2_end_total_tM9)
      
      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_tM9,
        value_table:"M9"
      })
      ws.getCell("N"+countM9).value = Number(this.sum_plug1_2_end_tM9)
      ws.getCell("N"+countM9).numFmt = '0.00';
      countM9++
         }
         if(this.sum_ttl_marker_tM9 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_tM9
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_tM9 = this.summary_plug12_tM9 / this.sum_ttl_marker_tM9
        if(this.sum_plug12_tM9 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_tM9
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_tM9 =  this.summary_widthloss_tM9 / this.sum_ttl_marker_tM9
        if(this.sum_per_width_loss_tM9 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_tM9
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_tM9 =  this.summary_end1_2_tM9/ this.sum_ttl_marker_tM9 //M
      if(this.sum_end1_2_tM9 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_tM9
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_tM9 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_tM9
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_tM9 = this.sum_endless_length_yd_tM9 / this.sum_ttl_marker_tM9 *100 //P
      if(this.sum_avg_end_tM9 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_tM9
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_tM9 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_tM9
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_tM9 = this.sum_spliceloss_tM9 / this.sum_ttl_marker_tM9 *100
      if(this.sum_splicelossper_tM9 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_tM9
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_tM9 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_tM9
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_tM9 = this.sum_totalcutout_tM9 / this.sum_ttl_marker_tM9 *100
      if(this.sum_totalcutoutper_tM9 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_tM9
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_tM9 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_tM9
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_tM9 = this.sum_rement_tM9 / this.sum_ttl_marker_tM9 *100
  if(this.sum_rement_per_tM9 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_tM9
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_tM9
    if(this.sum_totallenthloss_tM9 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_tM9
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspertM9 = this.sum_totallenthlosstM9 / this.sum_ttl_marker_tM9
    if(this.sum_totallenthlosspertM9 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspertM9
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }


     ws.getCell("A"+countxM9).value = "Total Table M9"  
      ws.mergeCells("A"+countxM9+":"+"I"+countxM9)
      
      if(this.sum_ttl_marker_tM9 > 0 && isNaN(this.sum_ttl_marker_tM9)==false && isFinite(this.sum_ttl_marker_tM9)==true){
      ws.getCell("J"+countxM9).value = Number(this.sum_ttl_marker_tM9)
      ws.getCell("J"+countxM9).numFmt = '0.00';
      }else{
      ws.getCell("J"+countxM9).value = 0
      ws.getCell("J"+countxM9).numFmt = '0.00';
      }

      if(this.sum_plug12_tM9 > 0 && isNaN(this.sum_plug12_tM9)==false && isFinite(this.sum_plug12_tM9)==true){
      ws.getCell("K"+countxM9).value = Number(this.sum_plug12_tM9)
      ws.getCell("K"+countxM9).numFmt = '0.00';
      }else{
      ws.getCell("K"+countxM9).value = 0
      ws.getCell("K"+countxM9).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_tM9 > 0 && isNaN(this.sum_per_width_loss_tM9)==false && isFinite(this.sum_per_width_loss_tM9)==true){
      ws.getCell("L"+countxM9).value = this.sum_per_width_loss_tM9/100
      ws.getCell("L"+countxM9).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countxM9).value = 0.00/100
      ws.getCell("L"+countxM9).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_tM9 = 0
      this.total_result_plug1_2_tM9 = this.sum_all_plug1_2_end_total_tM9 / this.sum_ttl_marker_tM9

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_tM9
      })
      if(this.total_result_plug1_2_tM9 > 0 && isNaN(this.total_result_plug1_2_tM9)==false && isFinite(this.total_result_plug1_2_tM9)==true){
      ws.getCell("N"+countxM9).value = Number(this.total_result_plug1_2_tM9)
      ws.getCell("N"+countxM9).numFmt = '0.00';
      }else{
      ws.getCell("N"+countxM9).value = 0.00
      ws.getCell("N"+countxM9).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_tM9 > 0 && isNaN(this.sum_endless_length_yd_tM9)==false && isFinite(this.sum_endless_length_yd_tM9)==true){
      ws.getCell("O"+countxM9).value = Number(this.sum_endless_length_yd_tM9)
      ws.getCell("O"+countxM9).numFmt = '0.00';
      }else{
      ws.getCell("O"+countxM9).value = 0.00
      ws.getCell("O"+countxM9).numFmt = '0.00';
      }

      if(this.sum_avg_end_tM9 > 0 && isNaN(this.sum_avg_end_tM9)==false && isFinite(this.sum_avg_end_tM9)==true){
      ws.getCell("P"+countxM9).value = this.sum_avg_end_tM9/100
      ws.getCell("P"+countxM9).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countxM9).value = 0/100
      ws.getCell("P"+countxM9).numFmt = '0.00%'; 
      }
  
      if(this.sum_spliceloss_tM9 > 0 && isNaN(this.sum_spliceloss_tM9)==false && isFinite(this.sum_spliceloss_tM9)==true){
      ws.getCell("R"+countxM9).value = Number(this.sum_spliceloss_tM9)
      ws.getCell("R"+countxM9).numFmt = '0.00';
      }else{
       ws.getCell("R"+countxM9).value = 0.00
      ws.getCell("R"+countxM9).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_tM9 > 0 && isNaN(this.sum_splicelossper_tM9)==false && isFinite(this.sum_splicelossper_tM9)==true){
      ws.getCell("S"+countxM9).value = this.sum_splicelossper_tM9/100
      ws.getCell("S"+countxM9).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countxM9).value = 0/100
      ws.getCell("S"+countxM9).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_tM9 > 0 && isNaN(this.sum_totalcutout_tM9)==false && isFinite(this.sum_totalcutout_tM9)==true){
      ws.getCell("U"+countxM9).value = Number(this.sum_totalcutout_tM9)
      ws.getCell("U"+countxM9).numFmt = '0.00';
      }else{
       ws.getCell("U"+countxM9).value = 0.00
      ws.getCell("U"+countxM9).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_tM9 > 0 && isNaN(this.sum_totalcutoutper_tM9)==false && isFinite(this.sum_totalcutoutper_tM9)==true){
      ws.getCell("V"+countxM9).value = this.sum_totalcutoutper_tM9/100
      ws.getCell("V"+countxM9).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countxM9).value = 0.00/100
      ws.getCell("V"+countxM9).numFmt = '0.00%';
      }

       if(this.sum_rement_tM9 > 0 && isNaN(this.sum_rement_tM9)==false && isFinite(this.sum_rement_tM9)==true){
      ws.getCell("X"+countxM9).value = Number(this.sum_rement_tM9)
      ws.getCell("X"+countxM9).numFmt = '0.00';
       }else{
      ws.getCell("X"+countxM9).value = 0.00
      ws.getCell("X"+countxM9).numFmt = '0.00';  
       }

      if(this.sum_rement_per_tM9 > 0 && isNaN(this.sum_rement_per_tM9)==false && isFinite(this.sum_rement_per_tM9)==true){
      ws.getCell("Y"+countxM9).value =  this.sum_rement_per_tM9/100
      ws.getCell("Y"+countxM9).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countxM9).value =  0.00/100
      ws.getCell("Y"+countxM9).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_tM9 > 0 && isNaN(this.sum_totallenthloss_tM9)==false && isFinite(this.sum_totallenthloss_tM9)==true){
      ws.getCell("AA"+countxM9).value = Number(this.sum_totallenthloss_tM9)
      ws.getCell("AA"+countxM9).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countxM9).value = 0
      ws.getCell("AA"+countxM9).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_tM9 = this.sum_totallenthloss_tM9 / this.sum_ttl_marker_tM9 *100
      if(this.last_sum_totallenthloss_per_tM9 > 0 && isNaN(this.last_sum_totallenthloss_per_tM9)==false && isFinite(this.last_sum_totallenthloss_per_tM9)==true){
      ws.getCell("AB"+countxM9).value = this.last_sum_totallenthloss_per_tM9/100
      ws.getCell("AB"+countxM9).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countxM9).value = 0/100
      ws.getCell("AB"+countxM9).numFmt = '0.00%'; 
      }

      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countxM9).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
      }

      //M10

      if(this.rowexport_type_M9.length == 0){
       
         var countM10 = countM9

       }else{
         var countM10 = countxM9+1
       }

       

        var countxM10 = countM10 + 1 +Number(this.rowexport_type_M10.length)
      
      this.sum_ttl_marker_tM10 = 0
      this.sum_plug12_tM10 = 0
      this.sum_per_width_loss_tM10 = 0
      this.sum_end1_2_tM10 = 0
      this.sum_endless_length_yd_tM10= 0
      this.sum_avg_end_tM10 = 0
      this.sum_spliceloss_tM10 = 0
      this.sum_splicelossper_tM10 = 0
      this.sum_totalcutout_tM10 = 0
      this.sum_totalcutoutper_tM10 = 0
      this.sum_rement_tM10 = 0
      this.sum_rement_per_tM10 = 0
      this.sum_totallenthloss_tM10 =0 
      this.sum_totallenthlossper_tM10 = 0
      this.summary_widthloss_tM10 = 0
      this.summary_plug12_tM10 = 0
      this.summary_end1_2_tM10 = 0

      
     
      this.sum_all_plug1_2_end_total_tM10 = 0

     if(this.rowexport_type_M10.length < 1){
      
      count_row_grand++
      
      }else{
          for (var ax = 0; ax<this.rowexport_type_M10.length; ax++){ 
      ws.getCell("A"+countM10).value = this.rowexport_type_M10[ax].mu_date

      ws.getCell("B"+countM10).value = this.rowexport_type_M10[ax].gm_table

      ws.getCell("C"+countM10).value = Number(this.rowexport_type_M10[ax].gm_no_short)
      ws.getCell("C"+countM10).numFmt = '0.00';

      ws.getCell("D"+countM10).value = this.rowexport_type_M10[ax].customer

      ws.getCell("E"+countM10).value = this.rowexport_type_M10[ax].gm_no

      ws.getCell("F"+countM10).value = this.rowexport_type_M10[ax].fabric_type

      ws.getCell("G"+countM10).value = Number(this.rowexport_type_M10[ax].obs_width)
      ws.getCell("G"+countM10).numFmt = '0.00';

      ws.getCell("H"+countM10).value = Number(this.rowexport_type_M10[ax].width_inch)
      ws.getCell("H"+countM10).numFmt = '0.00';


      ws.getCell("I"+countM10).value = Number(this.rowexport_type_M10[ax].length_ydb)
      ws.getCell("I"+countM10).numFmt = '0.00';

      ws.getCell("J"+countM10).value = Number(this.rowexport_type_M10[ax].ttl_marker)
      ws.getCell("J"+countM10).numFmt = '0.00';

      ws.getCell("K"+countM10).value = Number(this.rowexport_type_M10[ax].plug12)
      ws.getCell("K"+countM10).numFmt = '0.00';

      ws.getCell("L"+countM10).value = this.rowexport_type_M10[ax].widthloss/100
      ws.getCell("L"+countM10).numFmt = '0.00%';

      ws.getCell("M"+countM10).value = this.rowexport_type_M10[ax].widthloss_code


      ws.getCell("O"+countM10).value = Number(this.rowexport_type_M10[ax].endless_length_yd)
      ws.getCell("O"+countM10).numFmt = '0.00';

      ws.getCell("P"+countM10).value = this.rowexport_type_M10[ax].avg_end/100
      ws.getCell("P"+countM10).numFmt = '0.00%';

      ws.getCell("Q"+countM10).value = this.rowexport_type_M10[ax].avg_end_code
      

      ws.getCell("R"+countM10).value = Number(this.rowexport_type_M10[ax].spliceloss)
      ws.getCell("R"+countM10).numFmt = '0.00';

      ws.getCell("S"+countM10).value = this.rowexport_type_M10[ax].splicelossper/100
      ws.getCell("S"+countM10).numFmt = '0.00%';

      ws.getCell("T"+countM10).value = this.rowexport_type_M10[ax].per_splice_loss_code
   

      ws.getCell("U"+countM10).value = Number(this.rowexport_type_M10[ax].totalcutout)
      ws.getCell("U"+countM10).numFmt = '0.00';

      ws.getCell("V"+countM10).value = this.rowexport_type_M10[ax].totalcutoutper/100
      ws.getCell("V"+countM10).numFmt = '0.00%';

      ws.getCell("W"+countM10).value = this.rowexport_type_M10[ax].total_cut_out_per_code

      ws.getCell("X"+countM10).value = Number(this.rowexport_type_M10[ax].rement)
      ws.getCell("X"+countM10).numFmt = '0.00';

      ws.getCell("Y"+countM10).value = this.rowexport_type_M10[ax].rement_per/100
      ws.getCell("Y"+countM10).numFmt = '0.00%';

      ws.getCell("Z"+countM10).value = this.rowexport_type_M10[ax].remnants_loss_code

      ws.getCell("AA"+countM10).value = Number(this.rowexport_type_M10[ax].totallenthloss)
      ws.getCell("AA"+countM10).numFmt = '0.00';

      ws.getCell("AB"+countM10).value = this.rowexport_type_M10[ax].totallenthlossper/100
      ws.getCell("AB"+countM10).numFmt = '0.00%';
      
      this.sum_plug12_tM10 = 0
      this.sum_ttl_marker_tM10 = Number(this.sum_ttl_marker_tM10) + Number(this.rowexport_type_M10[ax].ttl_marker)
      this.sum_plug12_tM10 = this.rowexport_type_M10[ax].plug12 * this.rowexport_type_M10[ax].ttl_marker
      this.sum_per_width_loss_tM10 = this.rowexport_type_M10[ax].widthloss * this.rowexport_type_M10[ax].ttl_marker
      this.sum_end1_2_tM10 = this.rowexport_type_M10[ax].end1_2 * this.rowexport_type_M10[ax].ttl_marker
      this.sum_endless_length_yd_tM10 = Number(this.sum_endless_length_yd_tM10) + Number(this.rowexport_type_M10[ax].endless_length_yd)
      this.sum_spliceloss_tM10 = Number(this.sum_spliceloss_tM10) + Number(this.rowexport_type_M10[ax].spliceloss)
      this.sum_totalcutout_tM10 = Number(this.sum_totalcutout_tM10) + Number(this.rowexport_type_M10[ax].totalcutout)
      this.sum_rement_tM10 = Number(this.sum_rement_tM10) + Number(this.rowexport_type_M10[ax].rement)
      this.summary_plug12_tM10 = Number(this.summary_plug12_tM10) + Number(this.sum_plug12_tM10)
      this.summary_widthloss_tM10 = Number(this.summary_widthloss_tM10) + Number(this.sum_per_width_loss_tM10)
      this.summary_end1_2_tM10 = Number(this.summary_end1_2_tM10) + Number(this.sum_end1_2_tM10)
      this.sum_totallenthloss_tM10_first = 0
      this.sum_totallenthloss_tM10_first = Number(this.rowexport_type_M10[ax].endless_length_yd) + Number(this.rowexport_type_M10[ax].spliceloss)
      +Number(this.rowexport_type_M10[ax].totalcutout) + Number(this.rowexport_type_M10[ax].rement)
      this.sum_totallenthloss_tM10 = Number(this.sum_totallenthloss_tM10) + Number(this.sum_totallenthloss_tM10_first)

      this.sum_totallenthloss_per_tM10_first = 0
      this.sum_totallenthloss_per_tM10_first = Number(this.rowexport_type_M10[ax].avg_end) + Number(this.rowexport_type_M10[ax].splicelossper)
      +Number(this.rowexport_type_M10[ax].totalcutoutper) + Number(this.rowexport_type_M10[ax].rement_per)
      this.sum_totallenthloss_per_tM10 = Number(this.sum_totallenthloss_tM10) + Number(this.sum_totallenthloss_per_tM10_first)
     this.sum_number_1_end_tM10 = 0 
      this.sum_number_1_end_tM10 = this.rowexport_type_M10[ax].endless_length_yd * 36
      this.sum_number_2_end_tM10 = 0
      this.sum_number_2_end_tM10 = this.rowexport_type_M10[ax].ttl_marker/this.rowexport_type_M10[ax].length_ydb
      this.sum_plug1_2_end_tM10 = 0
      this.sum_plug1_2_end_tM10 = this.sum_number_1_end_tM10/this.sum_number_2_end_tM10

      this.sum_plug1_2_end_total_tM10 = 0
     
      this.sum_plug1_2_end_total_tM10 = this.sum_plug1_2_end_tM10 * this.rowexport_type_M10[ax].ttl_marker
      this.sum_all_plug1_2_end_total_tM10 = Number(this.sum_all_plug1_2_end_total_tM10) + Number(this.sum_plug1_2_end_total_tM10)
      
      this.keep_end_plug1_2.push({
        value:this.sum_plug1_2_end_tM10,
        value_table:"M10"
      })
      ws.getCell("N"+countM10).value = Number(this.sum_plug1_2_end_tM10)
      ws.getCell("N"+countM10).numFmt = '0.00';
      countM10++
         }
         if(this.sum_ttl_marker_tM10 > 0){
         this.total_compar_ttl_marker.push({
           value:this.sum_ttl_marker_tM10
         })
         }else{
         this.total_compar_ttl_marker.push({
           value:0
         })
         }
         
this.sum_plug12_tM10 = this.summary_plug12_tM10 / this.sum_ttl_marker_tM10
        if(this.sum_plug12_tM10 > 0){
        this.total_compar_plug12.push({
          value:this.sum_plug12_tM10
        })
        }else{
           this.total_compar_plug12.push({
          value:0
        })
        }
this.sum_per_width_loss_tM10 =  this.summary_widthloss_tM10 / this.sum_ttl_marker_tM10
        if(this.sum_per_width_loss_tM10 > 0){
          this.total_compar_per_width_loss.push({
            value:this.sum_per_width_loss_tM10
          })
        }else{
          this.total_compar_per_width_loss.push({
            value:0
          })
        }
         

this.sum_end1_2_tM10 =  this.summary_end1_2_tM10/ this.sum_ttl_marker_tM10 //M
      if(this.sum_end1_2_tM10 > 0){
      this.total_compar_end1_2.push({
        value:this.sum_end1_2_tM10
      })
      }else{
        this.total_compar_end1_2.push({
        value:0
      })
      }

      
      if(this.sum_endless_length_yd_tM10 > 0){ //N
      this.total_compar_endless_length_yd.push({
        value:this.sum_endless_length_yd_tM10
      })
      }else{
      this.total_compar_endless_length_yd.push({
        value:0
      })
      }

      this.sum_avg_end_tM10 = this.sum_endless_length_yd_tM10 / this.sum_ttl_marker_tM10 *100 //P
      if(this.sum_avg_end_tM10 > 0){
        this.total_compar_avg_end.push({
          value:this.sum_avg_end_tM10
        })
      }else{
        this.total_compar_avg_end.push({
          value:0
        })
      }


      
      if(this.sum_spliceloss_tM10 > 0){   //R
        this.total_compar_spliceloss.push({
          value:this.sum_spliceloss_tM10
        })
      }else{
        this.total_compar_spliceloss.push({
          value:0
        })
      }

      this.sum_splicelossper_tM10 = this.sum_spliceloss_tM10 / this.sum_ttl_marker_tM10 *100
      if(this.sum_splicelossper_tM10 > 0){
        this.total_compar_splicelossper.push({
          value:this.sum_splicelossper_tM10
        })
      }else{
        this.total_compar_splicelossper.push({
          value:0
        })
      }

      if(this.sum_totalcutout_tM10 > 0 ){
        this.total_compar_totalcutout.push({
          value:this.sum_totalcutout_tM10
        })
      }else{
        this.total_compar_totalcutout.push({
          value:0
        })
      }

      this.sum_totalcutoutper_tM10 = this.sum_totalcutout_tM10 / this.sum_ttl_marker_tM10 *100
      if(this.sum_totalcutoutper_tM10 > 0){
        this.total_compar_totalcutoutper.push({
          value:this.sum_totalcutoutper_tM10
        })
      }else{
        this.total_compar_totalcutoutper.push({
          value:0
        })
      }

    
    if(this.sum_rement_tM10 > 0){
      this.total_compar_rement.push({
        value:this.sum_rement_tM10
      })
    }else{
      this.total_compar_rement.push({
        value:0
      })
    }

  this.sum_rement_per_tM10 = this.sum_rement_tM10 / this.sum_ttl_marker_tM10 *100
  if(this.sum_rement_per_tM10 > 0){
    this.total_compar_rement_per.push({
      value:this.sum_rement_per_tM10
    })
  }else{
    this.total_compar_rement_per.push({
      value:0
    })
  }

    this.sum_totallenthloss_tM10
    if(this.sum_totallenthloss_tM10 > 0){
      this.total_compar_totallenthloss.push({
        value:this.sum_totallenthloss_tM10
      })
    }else{
      this.total_compar_totallenthloss.push({
        value:0
      })
    }

    this.sum_totallenthlosspertM10 = this.sum_totallenthlosstM10 / this.sum_ttl_marker_tM10
    if(this.sum_totallenthlosspertM10 > 0){
        this.total_compar_totallenthlossper.push({
          value:this.sum_totallenthlosspertM10
        })
    }else{
            this.total_compar_totallenthlossper.push({
          value:0
        })
    }


     ws.getCell("A"+countxM10).value = "Total Table M10"  
      ws.mergeCells("A"+countxM10+":"+"I"+countxM10)
      
      if(this.sum_ttl_marker_tM10 > 0 && isNaN(this.sum_ttl_marker_tM10)==false && isFinite(this.sum_ttl_marker_tM10)==true){
      ws.getCell("J"+countxM10).value = Number(this.sum_ttl_marker_tM10)
      ws.getCell("J"+countxM10).numFmt = '0.00';
      }else{
      ws.getCell("J"+countxM10).value = 0
      ws.getCell("J"+countxM10).numFmt = '0.00';
      }

      if(this.sum_plug12_tM10 > 0 && isNaN(this.sum_plug12_tM10)==false && isFinite(this.sum_plug12_tM10)==true){
      ws.getCell("K"+countxM10).value = Number(this.sum_plug12_tM10)
      ws.getCell("K"+countxM10).numFmt = '0.00';
      }else{
      ws.getCell("K"+countxM10).value = 0
      ws.getCell("K"+countxM10).numFmt = '0.00';
      }


      if(this.sum_per_width_loss_tM10 > 0 && isNaN(this.sum_per_width_loss_tM10)==false && isFinite(this.sum_per_width_loss_tM10)==true){
      ws.getCell("L"+countxM10).value = this.sum_per_width_loss_tM10/100
      ws.getCell("L"+countxM10).numFmt = '0.00%';
      }else{
      ws.getCell("L"+countxM10).value = 0.00/100
      ws.getCell("L"+countxM10).numFmt = '0.00%'; 
      }

      this.total_result_plug1_2_tM10 = 0
      this.total_result_plug1_2_tM10 = this.sum_all_plug1_2_end_total_tM10 / this.sum_ttl_marker_tM10

     this.keep_end_plug1_2_table.push({
        value:this.total_result_plug1_2_tM10
      })
      if(this.total_result_plug1_2_tM10 > 0 && isNaN(this.total_result_plug1_2_tM10)==false && isFinite(this.total_result_plug1_2_tM10)==true){
      ws.getCell("N"+countxM10).value = Number(this.total_result_plug1_2_tM10)
      ws.getCell("N"+countxM10).numFmt = '0.00';
      }else{
      ws.getCell("N"+countxM10).value = 0.00
      ws.getCell("N"+countxM10).numFmt = '0.00';
      }

      if(this.sum_endless_length_yd_tM10 > 0 && isNaN(this.sum_endless_length_yd_tM10)==false && isFinite(this.sum_endless_length_yd_tM10)==true){
      ws.getCell("O"+countxM10).value = Number(this.sum_endless_length_yd_tM10)
      ws.getCell("O"+countxM10).numFmt = '0.00';
      }else{
      ws.getCell("O"+countxM10).value = 0.00
      ws.getCell("O"+countxM10).numFmt = '0.00';
      }

      if(this.sum_avg_end_tM10 > 0 && isNaN(this.sum_avg_end_tM10)==false && isFinite(this.sum_avg_end_tM10)==true){
      ws.getCell("P"+countxM10).value = this.sum_avg_end_tM10/100
      ws.getCell("P"+countxM10).numFmt = '0.00%';
      }else{
       ws.getCell("P"+countxM10).value = 0/100
      ws.getCell("P"+countxM10).numFmt = '0.00%'; 
      }

      if(this.sum_spliceloss_tM10 > 0 && isNaN(this.sum_spliceloss_tM10)==false && isFinite(this.sum_spliceloss_tM10)==true){
      ws.getCell("R"+countxM10).value = Number(this.sum_spliceloss_tM10)
      ws.getCell("R"+countxM10).numFmt = '0.00';
      }else{
       ws.getCell("R"+countxM10).value = 0.00
      ws.getCell("R"+countxM10).numFmt = '0.00'; 
      }

      if(this.sum_splicelossper_tM10 > 0 && isNaN(this.sum_splicelossper_tM10)==false && isFinite(this.sum_splicelossper_tM10)==true){
      ws.getCell("S"+countxM10).value = this.sum_splicelossper_tM10/100
      ws.getCell("S"+countxM10).numFmt = '0.00%';
      }else{
      ws.getCell("S"+countxM10).value = 0/100
      ws.getCell("S"+countxM10).numFmt = '0.00%'; 
      }

      if(this.sum_totalcutout_tM10 > 0 && isNaN(this.sum_totalcutout_tM10)==false && isFinite(this.sum_totalcutout_tM10)==true){
      ws.getCell("U"+countxM10).value = Number(this.sum_totalcutout_tM10)
      ws.getCell("U"+countxM10).numFmt = '0.00';
      }else{
       ws.getCell("U"+countxM10).value = 0.00
      ws.getCell("U"+countxM10).numFmt = '0.00'; 
      }

      if(this.sum_totalcutoutper_tM10 > 0 && isNaN(this.sum_totalcutoutper_tM10)==false && isFinite(this.sum_totalcutoutper_tM10)==true){
      ws.getCell("V"+countxM10).value = this.sum_totalcutoutper_tM10/100
      ws.getCell("V"+countxM10).numFmt = '0.00%';
      }else{
      ws.getCell("V"+countxM10).value = 0.00/100
      ws.getCell("V"+countxM10).numFmt = '0.00%';
      }

       if(this.sum_rement_tM10 > 0 && isNaN(this.sum_rement_tM10)==false && isFinite(this.sum_rement_tM10)==true){
      ws.getCell("X"+countxM10).value = Number(this.sum_rement_tM10)
      ws.getCell("X"+countxM10).numFmt = '0.00';
       }else{
      ws.getCell("X"+countxM10).value = 0.00
      ws.getCell("X"+countxM10).numFmt = '0.00';  
       }

      if(this.sum_rement_per_tM10 > 0 && isNaN(this.sum_rement_per_tM10)==false && isFinite(this.sum_rement_per_tM10)==true){
      ws.getCell("Y"+countxM10).value =  this.sum_rement_per_tM10/100
      ws.getCell("Y"+countxM10).numFmt = '0.00%';
      }else{
      ws.getCell("Y"+countxM10).value =  0.00/100
      ws.getCell("Y"+countxM10).numFmt = '0.00%';
      }

      if(this.sum_totallenthloss_tM10 > 0 && isNaN(this.sum_totallenthloss_tM10)==false && isFinite(this.sum_totallenthloss_tM10)==true){
      ws.getCell("AA"+countxM10).value = Number(this.sum_totallenthloss_tM10)
      ws.getCell("AA"+countxM10).numFmt = '0.00';
      }else{
        ws.getCell("AA"+countxM10).value = 0
      ws.getCell("AA"+countxM10).numFmt = '0.00';
      }


      this.last_sum_totallenthloss_per_tM10 = this.sum_totallenthloss_tM10 / this.sum_ttl_marker_tM10 *100
      if(this.last_sum_totallenthloss_per_tM10 > 0 && isNaN(this.last_sum_totallenthloss_per_tM10)==false && isFinite(this.last_sum_totallenthloss_per_tM10)==true){
      ws.getCell("AB"+countxM10).value = this.last_sum_totallenthloss_per_tM10/100
      ws.getCell("AB"+countxM10).numFmt = '0.00%';
      }else{
       ws.getCell("AB"+countxM10).value = 0/100
      ws.getCell("AB"+countxM10).numFmt = '0.00%'; 
      }
 
      for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countxM10).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
       }
       
       
       var countLast = (countxM10 + 1 ) -2
    
       ws.getCell("A"+countLast).value = "Grand Total"
       ws.mergeCells("A"+ countLast+":"+"I"+countLast)

       for(var ax = 0; ax<this.column_second.length; ax++){
      ws.getCell(this.column_second[ax].col_name+countLast).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }
    

      
       
       this.grand_total_ttl_marker = 0
       this.grand_total_ttl_marker = Number(this.sum_ttl_marker_t1)
       +Number(this.sum_ttl_marker_t2)
       +Number(this.sum_ttl_marker_t3)
       +Number(this.sum_ttl_marker_t4)
       +Number(this.sum_ttl_marker_t5)
       +Number(this.sum_ttl_marker_t6)
       +Number(this.sum_ttl_marker_t7)
       +Number(this.sum_ttl_marker_t8)
       +Number(this.sum_ttl_marker_t9)
       +Number(this.sum_ttl_marker_t10)
       +Number(this.sum_ttl_marker_t11)
       +Number(this.sum_ttl_marker_t12)
       +Number(this.sum_ttl_marker_t13)
       +Number(this.sum_ttl_marker_t14)
       +Number(this.sum_ttl_marker_t15)
       +Number(this.sum_ttl_marker_t16)
       +Number(this.sum_ttl_marker_t17)
       +Number(this.sum_ttl_marker_t18)
       +Number(this.sum_ttl_marker_t19)
       +Number(this.sum_ttl_marker_t20)
       +Number(this.sum_ttl_marker_tM1)
       +Number(this.sum_ttl_marker_tM2)
       +Number(this.sum_ttl_marker_tM3)
       +Number(this.sum_ttl_marker_tM4)
       +Number(this.sum_ttl_marker_tM5)
       +Number(this.sum_ttl_marker_tM6)
       +Number(this.sum_ttl_marker_tM7)
       +Number(this.sum_ttl_marker_tM8)
       +Number(this.sum_ttl_marker_tM9)
       +Number(this.sum_ttl_marker_tM10)
       

       ws.getCell("J"+countLast).value = Number(this.grand_total_ttl_marker)
       ws.getCell("J"+countLast).numFmt = '#,##0.00';
       
       this.grand_total_plug1_2 = this.result_sum_ttl_width / this.grand_total_ttl_marker

       ws.getCell("K"+countLast).value = Number(this.grand_total_plug1_2)
       ws.getCell("K"+countLast).numFmt = '0.00';

        this.grand_total_width_per = this.result_sum_ttl_width_per  / this.grand_total_ttl_marker
       ws.getCell("L"+countLast).value = this.grand_total_width_per/100
       ws.getCell("L"+countLast).numFmt = '0.00%';
      this.sum_all_plug1_2_end = 0
      console.log(this.rowexport.length)
      console.log(this.keep_end_plug1_2)
      for(var ax = 0; ax<this.rowexport.length; ax++){
        
        this.sum_result_plug1_2_end = 0
        this.sum_result_plug1_2_end = this.keep_end_plug1_2[ax].value * this.rowexport[ax].ttl_marker

        this.sum_all_plug1_2_end = Number(this.sum_all_plug1_2_end ) + Number(this.sum_result_plug1_2_end)
      }
      this.last_result_plug1_2_end = 0
      this.last_result_plug1_2_end = this.sum_all_plug1_2_end / this.grand_total_ttl_marker
      ws.getCell("N"+countLast).value = Number(this.last_result_plug1_2_end)
       ws.getCell("N"+countLast).numFmt = '0.00';

      this.grand_total_end_length =0
       this.grand_total_end_length = Number(this.sum_endless_length_yd_t1)
       +Number(this.sum_endless_length_yd_t2)
       +Number(this.sum_endless_length_yd_t3)
       +Number(this.sum_endless_length_yd_t4)
       +Number(this.sum_endless_length_yd_t5)
       +Number(this.sum_endless_length_yd_t6)
       +Number(this.sum_endless_length_yd_t7)
       +Number(this.sum_endless_length_yd_t8)
       +Number(this.sum_endless_length_yd_t9)
       +Number(this.sum_endless_length_yd_t10)
       +Number(this.sum_endless_length_yd_t11)
       +Number(this.sum_endless_length_yd_t12)
       +Number(this.sum_endless_length_yd_t13)
       +Number(this.sum_endless_length_yd_t14)
       +Number(this.sum_endless_length_yd_t15)
       +Number(this.sum_endless_length_yd_t16)
       +Number(this.sum_endless_length_yd_t17)
       +Number(this.sum_endless_length_yd_t18)
       +Number(this.sum_endless_length_yd_t19)
       +Number(this.sum_endless_length_yd_t20)
       +Number(this.sum_endless_length_yd_tM1)
       +Number(this.sum_endless_length_yd_tM2)
       +Number(this.sum_endless_length_yd_tM3)
       +Number(this.sum_endless_length_yd_tM4)
       +Number(this.sum_endless_length_yd_tM5)
       +Number(this.sum_endless_length_yd_tM6)
       +Number(this.sum_endless_length_yd_tM7)
       +Number(this.sum_endless_length_yd_tM8)
       +Number(this.sum_endless_length_yd_tM9)
       +Number(this.sum_endless_length_yd_tM10)
       

       ws.getCell("O"+countLast).value = Number(this.grand_total_end_length)
       ws.getCell("O"+countLast).numFmt = '0.00';

       this.grand_total_end_per = 0
       this.grand_total_end_per = (this.grand_total_end_length / this.grand_total_ttl_marker)*100

       ws.getCell("P"+countLast).value = this.grand_total_end_per/100
       ws.getCell("P"+countLast).numFmt = '0.00%';

       for(var ax = countLast+1; ax<countLast+50; ax++){
         for(var ay = 0; ay<this.column_second.length; ay++){
        ws.getCell(this.column_second[ay].col_name+ax).border = {
    top: {style:'none', color: {argb:'FFFFFFFF'}},
    left: {style:'none', color: {argb:'FFFFFFF'}},
    bottom: {style:'none', color: {argb:'FFFFFFF'}},
    right: {style:'none', color: {argb:'FFFFFFF'}}
        };
    }

      }


     this.grand_total_splice_length = 0

     this.grand_total_splice_length = Number(this.sum_spliceloss_t1)
       +Number(this.sum_spliceloss_t2)
       +Number(this.sum_spliceloss_t3)
       +Number(this.sum_spliceloss_t4)
       +Number(this.sum_spliceloss_t5)
       +Number(this.sum_spliceloss_t6)
       +Number(this.sum_spliceloss_t7)
       +Number(this.sum_spliceloss_t8)
       +Number(this.sum_spliceloss_t9)
       +Number(this.sum_spliceloss_t10)
       +Number(this.sum_spliceloss_t11)
       +Number(this.sum_spliceloss_t12)
       +Number(this.sum_spliceloss_t13)
       +Number(this.sum_spliceloss_t14)
       +Number(this.sum_spliceloss_t15)
       +Number(this.sum_spliceloss_t16)
       +Number(this.sum_spliceloss_t17)
       +Number(this.sum_spliceloss_t18)
       +Number(this.sum_spliceloss_t19)
       +Number(this.sum_spliceloss_t20)
       +Number(this.sum_spliceloss_tM1)
       +Number(this.sum_spliceloss_tM2)
       +Number(this.sum_spliceloss_tM3)
       +Number(this.sum_spliceloss_tM4)
       +Number(this.sum_spliceloss_tM5)
       +Number(this.sum_spliceloss_tM6)
       +Number(this.sum_spliceloss_tM7)
       +Number(this.sum_spliceloss_tM8)
       +Number(this.sum_spliceloss_tM9)
       +Number(this.sum_spliceloss_tM10)

       ws.getCell("R"+countLast).value = Number(this.grand_total_splice_length)
       ws.getCell("R"+countLast).numFmt = '0.00';
       this.grand_total_splice_per = 0
       this.grand_total_splice_per = (this.grand_total_splice_length / this.grand_total_ttl_marker)*100
       ws.getCell("S"+countLast).value = this.grand_total_splice_per/100
       ws.getCell("S"+countLast).numFmt = '0.00%';

       this.grand_total_cut_length = 0
       this.grand_total_cut_length = Number(this.sum_totalcutout_t1)
       +Number(this.sum_totalcutout_t2)
       +Number(this.sum_totalcutout_t3)
       +Number(this.sum_totalcutout_t4)
       +Number(this.sum_totalcutout_t5)
       +Number(this.sum_totalcutout_t6)
       +Number(this.sum_totalcutout_t7)
       +Number(this.sum_totalcutout_t8)
       +Number(this.sum_totalcutout_t9)
       +Number(this.sum_totalcutout_t10)
       +Number(this.sum_totalcutout_t11)
       +Number(this.sum_totalcutout_t12)
       +Number(this.sum_totalcutout_t13)
       +Number(this.sum_totalcutout_t14)
       +Number(this.sum_totalcutout_t15)
       +Number(this.sum_totalcutout_t16)
       +Number(this.sum_totalcutout_t17)
       +Number(this.sum_totalcutout_t18)
       +Number(this.sum_totalcutout_t19)
       +Number(this.sum_totalcutout_t20)
       +Number(this.sum_totalcutout_tM1)
       +Number(this.sum_totalcutout_tM2)
       +Number(this.sum_totalcutout_tM3)
       +Number(this.sum_totalcutout_tM4)
       +Number(this.sum_totalcutout_tM5)
       +Number(this.sum_totalcutout_tM6)
       +Number(this.sum_totalcutout_tM7)
       +Number(this.sum_totalcutout_tM8)
       +Number(this.sum_totalcutout_tM9)
       +Number(this.sum_totalcutout_tM10)

       ws.getCell("U"+countLast).value = Number(this.grand_total_cut_length)
       ws.getCell("U"+countLast).numFmt = '0.00';

       this.grand_total_cut_per = 0
       this.grand_total_cut_per = (this.grand_total_cut_length / this.grand_total_ttl_marker)*100
       ws.getCell("V"+countLast).value = this.grand_total_cut_per/100
       ws.getCell("V"+countLast).numFmt = '0.00%';

        this.grand_total_rem_length = 0
       this.grand_total_rem_length = Number(this.sum_rement_t1)
       +Number(this.sum_rement_t2)
       +Number(this.sum_rement_t3)
       +Number(this.sum_rement_t4)
       +Number(this.sum_rement_t5)
       +Number(this.sum_rement_t6)
       +Number(this.sum_rement_t7)
       +Number(this.sum_rement_t8)
       +Number(this.sum_rement_t9)
       +Number(this.sum_rement_t10)
       +Number(this.sum_rement_t11)
       +Number(this.sum_rement_t12)
       +Number(this.sum_rement_t13)
       +Number(this.sum_rement_t14)
       +Number(this.sum_rement_t15)
       +Number(this.sum_rement_t16)
       +Number(this.sum_rement_t17)
       +Number(this.sum_rement_t18)
       +Number(this.sum_rement_t19)
       +Number(this.sum_rement_t20)
       +Number(this.sum_rement_tM1)
       +Number(this.sum_rement_tM2)
       +Number(this.sum_rement_tM3)
       +Number(this.sum_rement_tM4)
       +Number(this.sum_rement_tM5)
       +Number(this.sum_rement_tM6)
       +Number(this.sum_rement_tM7)
       +Number(this.sum_rement_tM8)
       +Number(this.sum_rement_tM9)
       +Number(this.sum_rement_tM10)
        
        ws.getCell("X"+countLast).value = Number(this.grand_total_rem_length)
       ws.getCell("X"+countLast).numFmt = '0.00';

       this.grand_total_rem_per = 0
       this.grand_total_rem_per = (this.grand_total_rem_length / this.grand_total_ttl_marker)*100
       ws.getCell("Y"+countLast).value = this.grand_total_rem_per/100
       ws.getCell("Y"+countLast).numFmt = '0.00%';

        this.grand_total_length_loss = 0
        this.grand_total_length_loss = Number(this.sum_totallenthloss_t1)
       +Number(this.sum_totallenthloss_t2)
       +Number(this.sum_totallenthloss_t3)
       +Number(this.sum_totallenthloss_t4)
       +Number(this.sum_totallenthloss_t5)
       +Number(this.sum_totallenthloss_t6)
       +Number(this.sum_totallenthloss_t7)
       +Number(this.sum_totallenthloss_t8)
       +Number(this.sum_totallenthloss_t9)
       +Number(this.sum_totallenthloss_t10)
       +Number(this.sum_totallenthloss_t11)
       +Number(this.sum_totallenthloss_t12)
       +Number(this.sum_totallenthloss_t13)
       +Number(this.sum_totallenthloss_t14)
       +Number(this.sum_totallenthloss_t15)
       +Number(this.sum_totallenthloss_t16)
       +Number(this.sum_totallenthloss_t17)
       +Number(this.sum_totallenthloss_t18)
       +Number(this.sum_totallenthloss_t19)
       +Number(this.sum_totallenthloss_t20)
       +Number(this.sum_totallenthloss_tM1)
       +Number(this.sum_totallenthloss_tM2)
       +Number(this.sum_totallenthloss_tM3)
       +Number(this.sum_totallenthloss_tM4)
       +Number(this.sum_totallenthloss_tM5)
       +Number(this.sum_totallenthloss_tM6)
       +Number(this.sum_totallenthloss_tM7)
       +Number(this.sum_totallenthloss_tM8)
       +Number(this.sum_totallenthloss_tM9)
       +Number(this.sum_totallenthloss_tM10)

       ws.getCell("AA"+countLast).value = Number(this.grand_total_length_loss)
       ws.getCell("AA"+countLast).numFmt = '0.00';

       this.grand_total_loss_per = 0
       this.grand_total_loss_per = (this.grand_total_length_loss / this.grand_total_ttl_marker)*100
       ws.getCell("AB"+countLast).value = this.grand_total_loss_per/100
       ws.getCell("AB"+countLast).numFmt = '0.00%';
     
   //sheet3
   ws1.columns = [
        { key: "A", width: 12 },
        { key: "B", width: 12 },
        { key: "C", width: 12 },
        { key: "D", width: 12 },
        { key: "E", width: 0.1 },
        { key: "F", width: 12 },
        { key: "G", width: 12 },
        { key: "H", width: 12 },
        { key: "I", width: 12 },
        { key: "J", width: 12 },
        { key: "K", width: 12 },
        { key: "L", width: 15 },
        { key: "M", width: 15 },
        { key: "N", width: 15},
        { key: "O", width: 15},
        { key: "P", width: 15 },
        
      ];


   for(var ax = 0; ax<this.column_third.length; ax++){
         for(var bz = 1; bz<37;  bz++){
        ws1.getCell(this.column_main[ax].col_name+bz).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
      
       ws1.getCell(this.column_main[ax].col_name+bz).alignment = {
          horizontal: "center",
          vertical: "middle",
        };

        ws1.getCell(this.column_main[ax].col_name+bz).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      }
    }
      ws1.getCell("A1").value = "Spread Loss Performance Compare Analysis :"+this.org;
      ws1.mergeCells("A1:D1");

      ws1.getCell("A3").value = "Table Group"
      ws1.mergeCells("A3:A5");

      ws1.getCell("B3").value = "Ttl Marker"
      ws1.mergeCells("B3:B4");

      ws1.getCell("B5").value = "Length(yd)"

      ws1.getCell("C3").value ="Width Loss"
      ws1.mergeCells("C3:D3");

      ws1.getCell("C4").value = "Plug+12"

      ws1.getCell("C5").value = "(Inch)"

      ws1.getCell("D4").value = "%of"

      ws1.getCell("D5").value = "Width Loss"
      
      ws1.getCell("F3").value ="End Loss"
      ws1.mergeCells("F3:H3");

      ws1.getCell("F4").value = "Plug+12"

      ws1.getCell("F5").value = "(Inch)"

      ws1.getCell("G4").value = "Length"

      ws1.getCell("G5").value = "(YD)"

      ws1.getCell("H4").value = "%of"

      ws1.getCell("H5").value = "End"

      ws1.getCell("I3").value ="Splice Loss"
      ws1.mergeCells("I3:J3");

      ws1.getCell("I4").value = "Length"

      ws1.getCell("I5").value = "(YD)"

      ws1.getCell("J4").value = "%of"

      ws1.getCell("J5").value = "Splice"

      ws1.getCell("K3").value ="Cut Out Loss"
      ws1.mergeCells("K3:L3");

      ws1.getCell("K4").value = "Length"

      ws1.getCell("K5").value = "(YD)"

      ws1.getCell("L4").value = "%of"

      ws1.getCell("L5").value = "Cut Out Loss"

      ws1.getCell("M3").value ="Remnants Loss"
      ws1.mergeCells("M3:N3");

      ws1.getCell("M4").value = "Length"

      ws1.getCell("M5").value = "(YD)"

      ws1.getCell("N4").value = "%of"

      ws1.getCell("N5").value = "Remnants"

      ws1.getCell("O3").value ="Total Length Loss"
      ws1.mergeCells("O3:P3");

      ws1.getCell("O4").value = "Length"

      ws1.getCell("O5").value = "(YD)"

      ws1.getCell("P4").value = "%of"

      ws1.getCell("P5").value = "Total Length"

      ws1.getCell("A6").value = "Table A1"
      ws1.getCell("A7").value = "Table A2"
      ws1.getCell("A8").value = "Table A3"
      ws1.getCell("A9").value = "Table A4"
      ws1.getCell("A10").value = "Table A5"
      ws1.getCell("A11").value = "Table A6"
      ws1.getCell("A12").value = "Table A7"
      ws1.getCell("A13").value = "Table A8"
      ws1.getCell("A14").value = "Table A9"
      ws1.getCell("A15").value = "Table A10"
      ws1.getCell("A16").value = "Table A11"
      ws1.getCell("A17").value = "Table A12"
      ws1.getCell("A18").value = "Table A13"
      ws1.getCell("A19").value = "Table A14"
      ws1.getCell("A20").value = "Table A15"
      ws1.getCell("A21").value = "Table A16"
      ws1.getCell("A22").value = "Table A17"
      ws1.getCell("A23").value = "Table A18"
      ws1.getCell("A24").value = "Table A19"
      ws1.getCell("A25").value = "Table A20"
      ws1.getCell("A26").value = "Table M1"
      ws1.getCell("A27").value = "Table M2"
      ws1.getCell("A28").value = "Table M3"
      ws1.getCell("A29").value = "Table M4"
      ws1.getCell("A30").value = "Table M5"
      ws1.getCell("A31").value = "Table M6"
      ws1.getCell("A32").value = "Table M7"
      ws1.getCell("A33").value = "Table M8"
      ws1.getCell("A34").value = "Table M9"
      ws1.getCell("A35").value = "Table M10"
      

      ws1.getCell("A36").value = "Total"
       ws1.getCell("A36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      console.log(this.keep_end_plug1_2_table)
      this.all_sum_total_per = 0
      for(var ax = 0; ax<this.total_compar_ttl_marker.length; ax++){
        ws1.getCell("B"+[ax+6]).value = Number(this.total_compar_ttl_marker[ax].value)
        ws1.getCell("B"+[ax+6]).numFmt = '#,##0.00';

        ws1.getCell("C"+[ax+6]).value = Number(this.total_compar_plug12[ax].value)
        ws1.getCell("C"+[ax+6]).numFmt = '0.00';

        ws1.getCell("D"+[ax+6]).value = this.total_compar_per_width_loss[ax].value/100
        ws1.getCell("D"+[ax+6]).numFmt = '0.00%';
        if(this.keep_end_plug1_2_table[ax].value > 0 && isNaN(this.keep_end_plug1_2_table[ax].value)==false && isFinite(this.keep_end_plug1_2_table[ax].value)==true){
        ws1.getCell("F"+[ax+6]).value = Number(this.keep_end_plug1_2_table[ax].value)
        ws1.getCell("F"+[ax+6]).numFmt = '0.00';
        }else{
          ws1.getCell("F"+[ax+6]).value = 0.00
        ws1.getCell("F"+[ax+6]).numFmt = '0.00';
        }
        ws1.getCell("G"+[ax+6]).value = Number(this.total_compar_endless_length_yd[ax].value)
        ws1.getCell("G"+[ax+6]).numFmt = '0.00';

        ws1.getCell("H"+[ax+6]).value = this.total_compar_avg_end[ax].value/100
        ws1.getCell("H"+[ax+6]).numFmt = '0.00%';

        ws1.getCell("I"+[ax+6]).value = Number(this.total_compar_spliceloss[ax].value)
        ws1.getCell("I"+[ax+6]).numFmt = '0.00';

        ws1.getCell("J"+[ax+6]).value = this.total_compar_splicelossper[ax].value/100
        ws1.getCell("J"+[ax+6]).numFmt = '0.00%';

        ws1.getCell("K"+[ax+6]).value = Number(this.total_compar_totalcutout[ax].value)
        ws1.getCell("K"+[ax+6]).numFmt = '0.00';

        ws1.getCell("L"+[ax+6]).value = this.total_compar_totalcutoutper[ax].value/100
        ws1.getCell("L"+[ax+6]).numFmt = '0.00%';

        ws1.getCell("M"+[ax+6]).value = Number(this.total_compar_rement[ax].value)
        ws1.getCell("M"+[ax+6]).numFmt = '0.00';

       
        ws1.getCell("N"+[ax+6]).value = this.total_compar_rement_per[ax].value/100
        ws1.getCell("N"+[ax+6]).numFmt = '0.00%';

        ws1.getCell("O"+[ax+6]).value = Number(this.total_compar_totallenthloss[ax].value)
        ws1.getCell("O"+[ax+6]).numFmt = '0.00';

       
        this.total_compar_all = 0
        this.total_compar_all = Number(this.total_compar_avg_end[ax].value) + Number(this.total_compar_splicelossper[ax].value)
        +Number(this.total_compar_totalcutoutper[ax].value) + Number(this.total_compar_rement_per[ax].value)

        this.all_sum_total_per = Number(this.all_sum_total_per) + Number(this.total_compar_all)
       
        ws1.getCell("P"+[ax+6]).value = this.total_compar_all/100
        ws1.getCell("P"+[ax+6]).numFmt = '0.00%';

        }

       
        this.sum_total_ttl_marker = 0 //sum ttl marker
        this.sum_total_end_loss = 0 //sum end loss
        this.sum_total_spliceloss = 0
        this.sum_total_cutoutloss = 0
        this.sum_total_rement = 0
        this.sum_total_totallength = 0
      
      for(var ax = 0; ax<this.total_compar_ttl_marker.length; ax++){
        this.sum_total_ttl_marker = Number(this.sum_total_ttl_marker) + Number(this.total_compar_ttl_marker[ax].value)
        this.sum_total_end_loss = Number(this.sum_total_end_loss) + Number(this.total_compar_endless_length_yd[ax].value)
        this.sum_total_spliceloss = Number(this.sum_total_spliceloss) + Number(this.total_compar_spliceloss[ax].value)
        this.sum_total_cutoutloss = Number(this.sum_total_cutoutloss) + Number(this.total_compar_totalcutout[ax].value)
        this.sum_total_rement = Number(this.sum_total_rement) + Number(this.total_compar_rement[ax].value)
        this.sum_total_totallength = Number(this.sum_total_totallength) + Number(this.total_compar_totallenthloss[ax].value)
      }

      if(this.sum_total_ttl_marker > 0){
      ws1.getCell("B36").value = Number(this.sum_total_ttl_marker)
      ws1.getCell("B36").numFmt = '0.00';
      ws1.getCell("B36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      }else{
      ws1.getCell("B36").value = Number(0)
      ws1.getCell("B36").numFmt = '0.00';
      ws1.getCell("B36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }

      if(this.sum_total_end_loss > 0){
      ws1.getCell("G36").value = Number(this.sum_total_end_loss)
      ws1.getCell("G36").numFmt = '0.00';
      ws1.getCell("G36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      
      }else{
      ws1.getCell("G36").value = Number(0)
      ws1.getCell("G36").numFmt = '0.00';
      ws1.getCell("G36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }

      if(this.sum_total_spliceloss > 0){
      ws1.getCell("I36").value = Number(this.sum_total_spliceloss)
      ws1.getCell("I36").numFmt = '0.00';
      ws1.getCell("I36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }else{
      ws1.getCell("I36").value = Number(0)
      ws1.getCell("I36").numFmt = '0.00';
      ws1.getCell("I36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }

      if(this.sum_total_cutoutloss > 0){
      ws1.getCell("K36").value = Number(this.sum_total_cutoutloss)
      ws1.getCell("K36").numFmt = '0.00';
      ws1.getCell("K36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }else{
      ws1.getCell("K36").value = Number(0)
      ws1.getCell("K36").numFmt = '0.00';
       ws1.getCell("K36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }

      if(this.sum_total_rement > 0){
      ws1.getCell("M36").value = Number(this.sum_total_rement)
      ws1.getCell("M36").numFmt = '0.00';
       ws1.getCell("M36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }else{
      ws1.getCell("M36").value = Number(0)
      ws1.getCell("M36").numFmt = '0.00';
      ws1.getCell("M36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }

      if(this.sum_total_totallength > 0){
      ws1.getCell("O36").value = Number(this.sum_total_totallength)
      ws1.getCell("O36").numFmt = '0.00';
      ws1.getCell("O36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }else{
      ws1.getCell("O36").value = Number(0)
      ws1.getCell("O36").numFmt = '0.00';
      ws1.getCell("O36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }


      


      this.sum_total_ttl_marker_wait = 0 //จำนวน C*B (N)
      this.sum_total_ttl_marker_wait_sec = 0
      this.total_sum_plug12 = 0
       for(var ax = 0; ax<this.total_compar_ttl_marker.length; ax++){
        this.sum_total_ttl_marker_wait = this.total_compar_endless_length_yd[ax].value * this.total_compar_ttl_marker[ax].value
        this.sum_total_ttl_marker_wait_sec = Number(this.sum_total_ttl_marker_wait_sec) + Number(this.sum_total_ttl_marker_wait)
      }  
      this.total_sum_plug12 = this.sum_total_ttl_marker_wait_sec / this.sum_total_ttl_marker
      
      if(this.grand_total_plug1_2 > 0){
      ws1.getCell("C36").value = Number(this.grand_total_plug1_2)
      ws1.getCell("C36").numFmt = '0.00';
      ws1.getCell("C36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }else{
      ws1.getCell("C36").value = Number(0)
      ws1.getCell("C36").numFmt = '0.00';
      ws1.getCell("C36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }

      this.sum_total_ttl_marker_wait_2 = 0 //จำนวน F*B (N)
      this.sum_total_ttl_marker_wait_sec_2 = 0
      this.total_sum_plug12_end = 0
       for(var ax = 0; ax<this.total_compar_ttl_marker.length; ax++){
        this.sum_total_ttl_marker_wait_2 = this.total_compar_end1_2[ax].value * this.total_compar_ttl_marker[ax].value
        this.sum_total_ttl_marker_wait_sec_2 = Number(this.sum_total_ttl_marker_wait_sec) + Number(this.sum_total_ttl_marker_wait_2)
      }  
      this.total_sum_plug12_end = this.sum_total_ttl_marker_wait_sec_2 / this.sum_total_ttl_marker
      
      if(this.total_sum_plug12_end > 0){
      ws1.getCell("F36").value = Number(this.last_result_plug1_2_end)
      ws1.getCell("F36").numFmt = '0.00';
      ws1.getCell("F36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }else{
      ws1.getCell("F36").value =  0
      ws1.getCell("F36").numFmt = '0.00';
      ws1.getCell("F36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }

      this.sum_total_ttl_marker_wait_3 = 0 //จำนวน D*B (N)
      this.sum_total_ttl_marker_wait_sec_3 = 0
      this.total_sum_width_loss_per = 0
       for(var ax = 0; ax<this.total_compar_ttl_marker.length; ax++){
        this.sum_total_ttl_marker_wait_3 = this.total_compar_per_width_loss[ax].value * this.total_compar_ttl_marker[ax].value

        this.sum_total_ttl_marker_wait_sec_3 = Number(this.sum_total_ttl_marker_wait_sec_3) + Number(this.sum_total_ttl_marker_wait_3)
      
      }  
      this.total_sum_width_loss_per = this.sum_total_ttl_marker_wait_sec_3 / this.sum_total_ttl_marker 


      if(this.total_sum_width_loss_per > 0){
      ws1.getCell("D36").value = this.total_sum_width_loss_per/100
      ws1.getCell("D36").numFmt = '0.00%';
      ws1.getCell("D36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }else{
      ws1.getCell("D36").value = 0/100
      ws1.getCell("D36").numFmt = '0.00%';
      ws1.getCell("D36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }


    

      if(this.grand_total_end_per > 0){
      ws1.getCell("H36").value = this.grand_total_end_per/100
      ws1.getCell("H36").numFmt = '0.00%';
      ws1.getCell("H36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }else{
      ws1.getCell("H36").value = 0/100
      ws1.getCell("H36").numFmt = '0.00%';
      ws1.getCell("H36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }


      if(this.grand_total_splice_per > 0){
      ws1.getCell("J36").value = this.grand_total_splice_per/100
      ws1.getCell("J36").numFmt = '0.00%';
      ws1.getCell("J36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }else{
      ws1.getCell("J36").value = 0/100
      ws1.getCell("J36").numFmt = '0.00%';
      ws1.getCell("J36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }



      if(this.grand_total_cut_per > 0){
      ws1.getCell("L36").value = this.grand_total_cut_per/100
      ws1.getCell("L36").numFmt = '0.00%';
      ws1.getCell("L36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }else{
      ws1.getCell("L36").value = 0/100
      ws1.getCell("L36").numFmt = '0.00%';
      ws1.getCell("L36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }

   

      if(this.grand_total_rem_per > 0){
      ws1.getCell("N36").value = this.grand_total_rem_per/100
      ws1.getCell("N36").numFmt = '0.0000%';
      ws1.getCell("N36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }else{
      ws1.getCell("N36").value = 0/100
      ws1.getCell("N36").numFmt = '0.00%';
      ws1.getCell("N36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }



       if(this.grand_total_loss_per > 0){
      ws1.getCell("P36").value = this.grand_total_loss_per/100
      ws1.getCell("P36").numFmt = '0.00%';
      ws1.getCell("P36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }else{
      ws1.getCell("P36").value = 0/100
      ws1.getCell("P36").numFmt = '0.00%';
      ws1.getCell("P36").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      }







      //ws2
     
      
      if(this.start_h !== ""){
       var startDate = moment(this.start_h,'DD-MM-YYYY')
       var endDate = moment(this.end_h,'DD-MM-YYYY')

       var first_date = new Date(startDate)

       var sec_date = new Date(this.end_h)
       
       this.fix_date = endDate.diff(startDate, 'days')  
       this.fix_week = this.fix_date/6
   
       this.fix_week = Math.round(this.fix_week)
    
       if(this.fix_date < 7){
         this.fix_week = 1
       }
        //round
     
      var round = 0
      
       for(var ax = 0; ax<this.fix_week; ax++){
         this.day_left = 0
        if(round == 0){
       this.firstNewDate = moment(first_date).format("YYYY/MM/DD");
       this.change_formar_first_date = moment(first_date).format("DD/MM");
       this.firstDate = this.firstNewDate

       this.next_date =  first_date.setDate(first_date.getDate() + 5);
       this.secDate =  moment(first_date).format("YYYY/MM/DD");
       this.change_formar_sec_date = moment(first_date).format("DD/MM");
       
       this.date_use_data.push({
         first_date : this.firstDate,
         sec_date : this.secDate,
         fmt_first_date:this.change_formar_first_date,
         fmt_sec_date:this.change_formar_sec_date
       })

       
       round++
       
      }else if(round > 0 && round < this.fix_week-1){
       this.next_date =  first_date.setDate(first_date.getDate() + 2);
       this.firstDate = moment(first_date).format("YYYY/MM/DD");
       this.change_formar_first_date = moment(first_date).format("DD/MM");

       this.next_date =  first_date.setDate(first_date.getDate() + 5);
       this.secDate =  moment(first_date).format("YYYY/MM/DD");
       this.change_formar_sec_date = moment(first_date).format("DD/MM");

       this.date_use_data.push({
         first_date : this.firstDate,
         sec_date : this.secDate,
         fmt_first_date:this.change_formar_first_date,
         fmt_sec_date:this.change_formar_sec_date
       })
       round++
     
  
     
       
    }else if(round == this.fix_week-1){
      this.day_left =  endDate.diff(this.secDate, 'days')
     
      this.next_date =  first_date.setDate(first_date.getDate() + 2);
       this.firstDate = moment(first_date).format("YYYY/MM/DD");
       this.change_formar_first_date = moment(first_date).format("DD/MM");
   
       this.next_date =  first_date.setDate(first_date.getDate() + 5);
       this.secDate =  moment(first_date).format("YYYY/MM/DD");
       this.change_formar_sec_date = moment(first_date).format("DD/MM");
       this.date_use_data.push({
         first_date : this.firstDate,
         sec_date : this.secDate,
         fmt_first_date:this.change_formar_first_date,
         fmt_sec_date:this.change_formar_sec_date
       })
       round++
    }
     
     }
     
      
    }else{
   
var myNewDate2 = new Date(this.first_fix_date);
         var count2 = 0;
for (var a1 = 0; a1 < 52; a1++) {
    
        
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
    }
      
      if(this.start_h !== ""){
      
      const params3 = new FormData();
       for(var ax = 0; ax<this.date_use_data.length; ax++){
         params3.append("start",this.date_use_data[ax].first_date)
         params3.append("end",this.date_use_data[ax].sec_date)
         params3.append("org",this.org)
        await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data_weekly_analy',
        data:params3,
       }).then((resp)=>{
        
          if (resp.data.data.length > 0) {
           this.rowexport_1_h=[]
          resp.data.data.forEach((e) => {
            this.rowexport_1_h.push(
              {
              gm_table: e.GM_TABLE,
              so_no_doc: e.SO_NO_DOC,
              ttl_marker: e.TTL_MARKER_YD,
              widthloss: e.WITH_LOSS,
              avg_end: e.AVG_END_LOSS,
              endless_length_yd: e.AVG_END_LOSS_YD,
              org:e.ORG,
              length_yd: e.LENGTH_YD,
              ttl_marker:e.TTL_MARKER_YD,
              spliceloss:e.SPLICE_LOSS_YD,
              totalcutout:e.TOTAL_CUTOUT_YD,
              rement:e.REMENT_LOSS_YD,
            })
          
          });
        
       this.sum_ttl_marker_t1_h = 0
       this.sum_width_end_t1_h = 0
       this.result_sum_width_end_t1_h = 0
       this.sum_end_loss_length_yd_t1 = 0
       this.result_end_loss_yd_t1 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A1"){
           
           this.sum_width_end_first_t1_h = 0
            
            this.sum_width_end_first_t1_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t1_h = Number(this.sum_width_end_t1_h) + Number(this.sum_width_end_first_t1_h)
            this.sum_ttl_marker_t1_h = Number(this.sum_ttl_marker_t1_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t1 = Number(this.sum_end_loss_length_yd_t1) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        } 
     
        this.result_sum_width_end_t1_h = this.sum_width_end_t1_h / this.sum_ttl_marker_t1_h
            if(isNaN(this.result_sum_width_end_t1_h)==false){
            this.row_result_t1.push({
              value:this.result_sum_width_end_t1_h
            }) 
            }else{
              this.row_result_t1.push({
              value:0
            })
            }

        this.result_end_loss_yd_t1 = this.sum_end_loss_length_yd_t1 / this.sum_ttl_marker_t1_h *100
            if(isNaN(this.result_end_loss_yd_t1)==false){
            this.row_result_endloss_t1.push({
              value:this.result_end_loss_yd_t1
            })
            }else{
              this.row_result_endloss_t1.push({
              value:0
            })
            }

        
            
            
      this.sum_ttl_marker_t2_h = 0
       this.sum_width_end_t2_h = 0
       this.result_sum_width_end_t2_h = 0
       this.sum_end_loss_length_yd_t2 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A2"){
           
           this.sum_width_end_first_t2_h = 0
            this.sum_width_end_first_t2_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t2_h = Number(this.sum_width_end_t2_h) + Number(this.sum_width_end_first_t2_h)
            this.sum_ttl_marker_t2_h = Number(this.sum_ttl_marker_t2_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
           
            this.sum_end_loss_length_yd_t2 = Number(this.sum_end_loss_length_yd_t2) + Number(this.rowexport_1_h[ax].endless_length_yd)
            }
        
        }  
       
        this.result_sum_width_end_t2_h = this.sum_width_end_t2_h / this.sum_ttl_marker_t2_h
            if(isNaN(this.result_sum_width_end_t2_h)==false){
            this.row_result_t2.push({
              value:this.result_sum_width_end_t2_h
            }) 
            }else{
              this.row_result_t2.push({
              value:0
            })
            }
        
       
        this.result_end_loss_yd_t2 = this.sum_end_loss_length_yd_t2 / this.sum_ttl_marker_t2_h*100
            if(isNaN(this.result_end_loss_yd_t2)==false){
            this.row_result_endloss_t2.push({
              value:this.result_end_loss_yd_t2
            })
            }else{
              this.row_result_endloss_t2.push({
              value:0
            })
            }

       this.sum_ttl_marker_t3_h = 0
       this.sum_width_end_t3_h = 0
       this.result_sum_width_end_t3_h = 0
        this.sum_end_loss_length_yd_t3 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A3"){
           
           this.sum_width_end_first_t3_h = 0
            this.sum_width_end_first_t3_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t3_h = Number(this.sum_width_end_t3_h) + Number(this.sum_width_end_first_t3_h)
            this.sum_ttl_marker_t3_h = Number(this.sum_ttl_marker_t3_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t3 = Number(this.sum_end_loss_length_yd_t3) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
   
         this.result_sum_width_end_t3_h = this.sum_width_end_t3_h / this.sum_ttl_marker_t3_h
   
            if(isNaN(this.result_sum_width_end_t3_h)==false){
            this.row_result_t3.push({
              value:this.result_sum_width_end_t3_h
            }) 
            }else{
              this.row_result_t3.push({
              value:0
            })
            }

        this.result_end_loss_yd_t3 = this.sum_end_loss_length_yd_t3 / this.sum_ttl_marker_t3_h*100
            if(isNaN(this.result_end_loss_yd_t3)==false){
            this.row_result_endloss_t3.push({
              value:this.result_end_loss_yd_t3
            })
            }else{
              this.row_result_endloss_t3.push({
              value:0
            })
            }

        
        this.sum_ttl_marker_t4_h = 0
       this.sum_width_end_t4_h = 0
       this.result_sum_width_end_t4_h = 0
        this.sum_end_loss_length_yd_t4 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A4"){
           
           this.sum_width_end_first_t4_h = 0
            this.sum_width_end_first_t4_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t4_h = Number(this.sum_width_end_t4_h) + Number(this.sum_width_end_first_t4_h)
            this.sum_ttl_marker_t4_h = Number(this.sum_ttl_marker_t4_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t4 = Number(this.sum_end_loss_length_yd_t4) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t4_h = this.sum_width_end_t4_h / this.sum_ttl_marker_t4_h
            if(isNaN(this.result_sum_width_end_t4_h)==false){
            this.row_result_t4.push({
              value:this.result_sum_width_end_t4_h
            }) 
            }else{
              this.row_result_t4.push({
              value:0
            })
            }

        this.result_end_loss_yd_t4 = this.sum_end_loss_length_yd_t4 / this.sum_ttl_marker_t4_h*100
            if(isNaN(this.result_end_loss_yd_t4)==false){
            this.row_result_endloss_t4.push({
              value:this.result_end_loss_yd_t4
            })
            }else{
              this.row_result_endloss_t4.push({
              value:0
            })
            }

       this.sum_ttl_marker_t5_h = 0
       this.sum_width_end_t5_h = 0
       this.result_sum_width_end_t5_h = 0
        this.sum_end_loss_length_yd_t5 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A5"){
           
           this.sum_width_end_first_t5_h = 0
            this.sum_width_end_first_t5_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t5_h = Number(this.sum_width_end_t5_h) + Number(this.sum_width_end_first_t5_h)
            this.sum_ttl_marker_t5_h = Number(this.sum_ttl_marker_t5_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t5 = Number(this.sum_end_loss_length_yd_t5) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t5_h = this.sum_width_end_t5_h / this.sum_ttl_marker_t5_h
            if(isNaN(this.result_sum_width_end_t5_h)==false){
            this.row_result_t5.push({
              value:this.result_sum_width_end_t5_h
            }) 
            }else{
              this.row_result_t5.push({
              value:0
            })
            }

            this.result_end_loss_yd_t5 = this.sum_end_loss_length_yd_t5 / this.sum_ttl_marker_t5_h*100
            if(isNaN(this.result_end_loss_yd_t5)==false){
            this.row_result_endloss_t5.push({
              value:this.result_end_loss_yd_t5
            })
            }else{
              this.row_result_endloss_t5.push({
              value:0
            })
            }

             this.sum_ttl_marker_t6_h = 0
       this.sum_width_end_t6_h = 0
       this.result_sum_width_end_t6_h = 0
        this.sum_end_loss_length_yd_t6 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A6"){
           
           this.sum_width_end_first_t6_h = 0
            this.sum_width_end_first_t6_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t6_h = Number(this.sum_width_end_t6_h) + Number(this.sum_width_end_first_t6_h)
            this.sum_ttl_marker_t6_h = Number(this.sum_ttl_marker_t6_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t6 = Number(this.sum_end_loss_length_yd_t6) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t6_h = this.sum_width_end_t6_h / this.sum_ttl_marker_t6_h
            if(isNaN(this.result_sum_width_end_t6_h)==false){
            this.row_result_t6.push({
              value:this.result_sum_width_end_t6_h
            }) 
            }else{
              this.row_result_t6.push({
              value:0
            })
            }

            this.result_end_loss_yd_t6 = this.sum_end_loss_length_yd_t6 / this.sum_ttl_marker_t6_h*100
            if(isNaN(this.result_end_loss_yd_t6)==false){
            this.row_result_endloss_t6.push({
              value:this.result_end_loss_yd_t6
            })
            }else{
              this.row_result_endloss_t6.push({
              value:0
            })
            }

       this.sum_ttl_marker_t7_h = 0
       this.sum_width_end_t7_h = 0
       this.result_sum_width_end_t7_h = 0
        this.sum_end_loss_length_yd_t7 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A7"){
           
           this.sum_width_end_first_t7_h = 0
            this.sum_width_end_first_t7_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t7_h = Number(this.sum_width_end_t7_h) + Number(this.sum_width_end_first_t7_h)
            this.sum_ttl_marker_t7_h = Number(this.sum_ttl_marker_t7_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t7 = Number(this.sum_end_loss_length_yd_t7) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t7_h = this.sum_width_end_t7_h / this.sum_ttl_marker_t7_h
            if(isNaN(this.result_sum_width_end_t7_h)==false){
            this.row_result_t7.push({
              value:this.result_sum_width_end_t7_h
            }) 
            }else{
              this.row_result_t7.push({
              value:0
            })
            }

          this.result_end_loss_yd_t7 = this.sum_end_loss_length_yd_t7 / this.sum_ttl_marker_t7_h*100
            if(isNaN(this.result_end_loss_yd_t7)==false){
            this.row_result_endloss_t7.push({
              value:this.result_end_loss_yd_t7
            })
            }else{
              this.row_result_endloss_t7.push({
              value:0
            })
            }
            
            this.sum_ttl_marker_t8_h = 0
       this.sum_width_end_t8_h = 0
       this.result_sum_width_end_t8_h = 0
        this.sum_end_loss_length_yd_t8 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A8"){
           
           this.sum_width_end_first_t8_h = 0
            this.sum_width_end_first_t8_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t8_h = Number(this.sum_width_end_t8_h) + Number(this.sum_width_end_first_t8_h)
            this.sum_ttl_marker_t8_h = Number(this.sum_ttl_marker_t8_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t8 = Number(this.sum_end_loss_length_yd_t8) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t8_h = this.sum_width_end_t8_h / this.sum_ttl_marker_t8_h
            if(isNaN(this.result_sum_width_end_t8_h)==false){
            this.row_result_t8.push({
              value:this.result_sum_width_end_t8_h
            }) 
            }else{
              this.row_result_t8.push({
              value:0
            })
            }

        this.result_end_loss_yd_t8 = this.sum_end_loss_length_yd_t8 / this.sum_ttl_marker_t8_h*100
            if(isNaN(this.result_end_loss_yd_t8)==false){
            this.row_result_endloss_t8.push({
              value:this.result_end_loss_yd_t8
            })
            }else{
              this.row_result_endloss_t8.push({
              value:0
            })
            }

             this.sum_ttl_marker_t9_h = 0
       this.sum_width_end_t9_h = 0
       this.result_sum_width_end_t9_h = 0
        this.sum_end_loss_length_yd_t9 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A9"){
           
           this.sum_width_end_first_t9_h = 0
            this.sum_width_end_first_t9_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t9_h = Number(this.sum_width_end_t9_h) + Number(this.sum_width_end_first_t9_h)
            this.sum_ttl_marker_t9_h = Number(this.sum_ttl_marker_t9_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t9 = Number(this.sum_end_loss_length_yd_t9) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t9_h = this.sum_width_end_t9_h / this.sum_ttl_marker_t9_h
            if(isNaN(this.result_sum_width_end_t9_h)==false){
            this.row_result_t9.push({
              value:this.result_sum_width_end_t9_h
            }) 
            }else{
              this.row_result_t9.push({
              value:0
            })
            }

        this.result_end_loss_yd_t9 = this.sum_end_loss_length_yd_t9 / this.sum_ttl_marker_t9_h*100
            if(isNaN(this.result_end_loss_yd_t9)==false){
            this.row_result_endloss_t9.push({
              value:this.result_end_loss_yd_t9
            })
            }else{
              this.row_result_endloss_t9.push({
              value:0
            })
            }

            this.sum_ttl_marker_t10_h = 0
       this.sum_width_end_t10_h = 0
       this.result_sum_width_end_t10_h = 0
        this.sum_end_loss_length_yd_t10 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A10"){
           
           this.sum_width_end_first_t10_h = 0
            this.sum_width_end_first_t10_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t10_h = Number(this.sum_width_end_t10_h) + Number(this.sum_width_end_first_t10_h)
            this.sum_ttl_marker_t10_h = Number(this.sum_ttl_marker_t10_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t10 = Number(this.sum_end_loss_length_yd_t10) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t10_h = this.sum_width_end_t10_h / this.sum_ttl_marker_t10_h
            if(isNaN(this.result_sum_width_end_t10_h)==false){
            this.row_result_t10.push({
              value:this.result_sum_width_end_t10_h
            }) 
            }else{
              this.row_result_t10.push({
              value:0
            })
            }

        this.result_end_loss_yd_t10 = this.sum_end_loss_length_yd_t10 / this.sum_ttl_marker_t10_h*100
            if(isNaN(this.result_end_loss_yd_t10)==false){
            this.row_result_endloss_t10.push({
              value:this.result_end_loss_yd_t10
            })
            }else{
              this.row_result_endloss_t10.push({
              value:0
            })
            }

            this.sum_ttl_marker_t11_h = 0
       this.sum_width_end_t11_h = 0
       this.result_sum_width_end_t11_h = 0
       this.sum_end_loss_length_yd_t11 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A11"){
           
           this.sum_width_end_first_t11_h = 0
            this.sum_width_end_first_t11_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t11_h = Number(this.sum_width_end_t11_h) + Number(this.sum_width_end_first_t11_h)
            this.sum_ttl_marker_t11_h = Number(this.sum_ttl_marker_t11_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t11 = Number(this.sum_end_loss_length_yd_t11) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t11_h = this.sum_width_end_t11_h / this.sum_ttl_marker_t11_h
            if(isNaN(this.result_sum_width_end_t11_h)==false){
            this.row_result_t11.push({
              value:this.result_sum_width_end_t11_h
            }) 
            }else{
              this.row_result_t11.push({
              value:0
            })
            }

        this.result_end_loss_yd_t11 = this.sum_end_loss_length_yd_t11 / this.sum_ttl_marker_t11_h*100
            if(isNaN(this.result_end_loss_yd_t11)==false){
            this.row_result_endloss_t11.push({
              value:this.result_end_loss_yd_t11
            })
            }else{
              this.row_result_endloss_t11.push({
              value:0
            })
            }

            this.sum_ttl_marker_t12_h = 0
       this.sum_width_end_t12_h = 0
       this.result_sum_width_end_t12_h = 0
       this.sum_end_loss_length_yd_t12 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A12"){
           
           this.sum_width_end_first_t12_h = 0
            this.sum_width_end_first_t12_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t12_h = Number(this.sum_width_end_t12_h) + Number(this.sum_width_end_first_t12_h)
            this.sum_ttl_marker_t12_h = Number(this.sum_ttl_marker_t12_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t12 = Number(this.sum_end_loss_length_yd_t12) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t12_h = this.sum_width_end_t12_h / this.sum_ttl_marker_t12_h
            if(isNaN(this.result_sum_width_end_t12_h)==false){
            this.row_result_t12.push({
              value:this.result_sum_width_end_t12_h
            }) 
            }else{
              this.row_result_t12.push({
              value:0
            })
            }

        
         this.result_end_loss_yd_t12 = this.sum_end_loss_length_yd_t12 / this.sum_ttl_marker_t12_h*100
            if(isNaN(this.result_end_loss_yd_t12)==false){
            this.row_result_endloss_t12.push({
              value:this.result_end_loss_yd_t12
            })
            }else{
              this.row_result_endloss_t12.push({
              value:0
            })
            }

            this.sum_ttl_marker_t13_h = 0
       this.sum_width_end_t13_h = 0
       this.result_sum_width_end_t13_h = 0
       this.sum_end_loss_length_yd_t13 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A13"){
           
           this.sum_width_end_first_t13_h = 0
            this.sum_width_end_first_t13_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t13_h = Number(this.sum_width_end_t13_h) + Number(this.sum_width_end_first_t13_h)
            this.sum_ttl_marker_t13_h = Number(this.sum_ttl_marker_t13_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t13 = Number(this.sum_end_loss_length_yd_t13) + Number(this.rowexport_1_h[ax].endless_length_yd)
            console.log(this.sum_width_end_t13_h)
            }
        
        }  
         this.result_sum_width_end_t13_h = this.sum_width_end_t13_h / this.sum_ttl_marker_t13_h
        
         
            if(isNaN(this.result_sum_width_end_t13_h)==false){
            this.row_result_t13.push({
              value:this.result_sum_width_end_t13_h
            }) 
            }else{
              this.row_result_t13.push({
              value:0
            })
            }

        
         this.result_end_loss_yd_t13 = this.sum_end_loss_length_yd_t13 / this.sum_ttl_marker_t13_h*100
            if(isNaN(this.result_end_loss_yd_t13)==false){
            this.row_result_endloss_t13.push({
              value:this.result_end_loss_yd_t13
            })
            }else{
              this.row_result_endloss_t13.push({
              value:0
            })
            }

            this.sum_ttl_marker_t14_h = 0
       this.sum_width_end_t14_h = 0
       this.result_sum_width_end_t14_h = 0
       this.sum_end_loss_length_yd_t14 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A14"){
           
           this.sum_width_end_first_t14_h = 0
            this.sum_width_end_first_t14_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t14_h = Number(this.sum_width_end_t14_h) + Number(this.sum_width_end_first_t14_h)
            this.sum_ttl_marker_t14_h = Number(this.sum_ttl_marker_t14_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t14 = Number(this.sum_end_loss_length_yd_t14) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t14_h = this.sum_width_end_t14_h / this.sum_ttl_marker_t14_h
            if(isNaN(this.result_sum_width_end_t14_h)==false){
            this.row_result_t14.push({
              value:this.result_sum_width_end_t14_h
            }) 
            }else{
              this.row_result_t14.push({
              value:0
            })
            }

        
         this.result_end_loss_yd_t14 = this.sum_end_loss_length_yd_t14 / this.sum_ttl_marker_t14_h*100
            if(isNaN(this.result_end_loss_yd_t14)==false){
            this.row_result_endloss_t14.push({
              value:this.result_end_loss_yd_t14
            })
            }else{
              this.row_result_endloss_t14.push({
              value:0
            })
            }

            this.sum_ttl_marker_t15_h = 0
       this.sum_width_end_t15_h = 0
       this.result_sum_width_end_t15_h = 0
       this.sum_end_loss_length_yd_t15 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A15"){
           
           this.sum_width_end_first_t15_h = 0
            this.sum_width_end_first_t15_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t15_h = Number(this.sum_width_end_t15_h) + Number(this.sum_width_end_first_t15_h)
            this.sum_ttl_marker_t15_h = Number(this.sum_ttl_marker_t15_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t15 = Number(this.sum_end_loss_length_yd_t15) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t15_h = this.sum_width_end_t15_h / this.sum_ttl_marker_t15_h
            if(isNaN(this.result_sum_width_end_t15_h)==false){
            this.row_result_t15.push({
              value:this.result_sum_width_end_t15_h
            }) 
            }else{
              this.row_result_t15.push({
              value:0
            })
            }

        
         this.result_end_loss_yd_t15 = this.sum_end_loss_length_yd_t15 / this.sum_ttl_marker_t15_h*100
            if(isNaN(this.result_end_loss_yd_t15)==false){
            this.row_result_endloss_t15.push({
              value:this.result_end_loss_yd_t15
            })
            }else{
              this.row_result_endloss_t15.push({
              value:0
            })
            }

            this.sum_ttl_marker_t16_h = 0
       this.sum_width_end_t16_h = 0
       this.result_sum_width_end_t16_h = 0
       this.sum_end_loss_length_yd_t16 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A16"){
           
           this.sum_width_end_first_t16_h = 0
            this.sum_width_end_first_t16_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t16_h = Number(this.sum_width_end_t16_h) + Number(this.sum_width_end_first_t16_h)
            this.sum_ttl_marker_t16_h = Number(this.sum_ttl_marker_t16_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t16 = Number(this.sum_end_loss_length_yd_t16) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t16_h = this.sum_width_end_t16_h / this.sum_ttl_marker_t16_h
            if(isNaN(this.result_sum_width_end_t16_h)==false){
            this.row_result_t16.push({
              value:this.result_sum_width_end_t16_h
            }) 
            }else{
              this.row_result_t16.push({
              value:0
            })
            }

        
         this.result_end_loss_yd_t16 = this.sum_end_loss_length_yd_t16 / this.sum_ttl_marker_t16_h*100
            if(isNaN(this.result_end_loss_yd_t16)==false){
            this.row_result_endloss_t16.push({
              value:this.result_end_loss_yd_t16
            })
            }else{
              this.row_result_endloss_t16.push({
              value:0
            })
            }

            this.sum_ttl_marker_t17_h = 0
       this.sum_width_end_t17_h = 0
       this.result_sum_width_end_t17_h = 0
       this.sum_end_loss_length_yd_t17 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A17"){
           
           this.sum_width_end_first_t17_h = 0
            this.sum_width_end_first_t17_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t17_h = Number(this.sum_width_end_t17_h) + Number(this.sum_width_end_first_t17_h)
            this.sum_ttl_marker_t17_h = Number(this.sum_ttl_marker_t17_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t17 = Number(this.sum_end_loss_length_yd_t17) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t17_h = this.sum_width_end_t17_h / this.sum_ttl_marker_t17_h
            if(isNaN(this.result_sum_width_end_t17_h)==false){
            this.row_result_t17.push({
              value:this.result_sum_width_end_t17_h
            }) 
            }else{
              this.row_result_t17.push({
              value:0
            })
            }

        
         this.result_end_loss_yd_t17 = this.sum_end_loss_length_yd_t17 / this.sum_ttl_marker_t17_h*100
            if(isNaN(this.result_end_loss_yd_t17)==false){
            this.row_result_endloss_t17.push({
              value:this.result_end_loss_yd_t17
            })
            }else{
              this.row_result_endloss_t17.push({
              value:0
            })
            }

            this.sum_ttl_marker_t18_h = 0
       this.sum_width_end_t18_h = 0
       this.result_sum_width_end_t18_h = 0
       this.sum_end_loss_length_yd_t18 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A18"){
           
           this.sum_width_end_first_t18_h = 0
            this.sum_width_end_first_t18_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t18_h = Number(this.sum_width_end_t18_h) + Number(this.sum_width_end_first_t18_h)
            this.sum_ttl_marker_t18_h = Number(this.sum_ttl_marker_t18_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t18 = Number(this.sum_end_loss_length_yd_t18) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t18_h = this.sum_width_end_t18_h / this.sum_ttl_marker_t18_h
            if(isNaN(this.result_sum_width_end_t18_h)==false){
            this.row_result_t18.push({
              value:this.result_sum_width_end_t18_h
            }) 
            }else{
              this.row_result_t18.push({
              value:0
            })
            }

        
         this.result_end_loss_yd_t18 = this.sum_end_loss_length_yd_t18 / this.sum_ttl_marker_t18_h*100
            if(isNaN(this.result_end_loss_yd_t18)==false){
            this.row_result_endloss_t18.push({
              value:this.result_end_loss_yd_t18
            })
            }else{
              this.row_result_endloss_t18.push({
              value:0
            })
            }

            this.sum_ttl_marker_t19_h = 0
       this.sum_width_end_t19_h = 0
       this.result_sum_width_end_t19_h = 0
       this.sum_end_loss_length_yd_t19 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A19"){
           
           this.sum_width_end_first_t19_h = 0
            this.sum_width_end_first_t19_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t19_h = Number(this.sum_width_end_t19_h) + Number(this.sum_width_end_first_t19_h)
            this.sum_ttl_marker_t19_h = Number(this.sum_ttl_marker_t19_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t19 = Number(this.sum_end_loss_length_yd_t19) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t19_h = this.sum_width_end_t19_h / this.sum_ttl_marker_t19_h
            if(isNaN(this.result_sum_width_end_t19_h)==false){
            this.row_result_t19.push({
              value:this.result_sum_width_end_t19_h
            }) 
            }else{
              this.row_result_t19.push({
              value:0
            })
            }

        
         this.result_end_loss_yd_t19 = this.sum_end_loss_length_yd_t19 / this.sum_ttl_marker_t19_h*100
            if(isNaN(this.result_end_loss_yd_t19)==false){
            this.row_result_endloss_t19.push({
              value:this.result_end_loss_yd_t19
            })
            }else{
              this.row_result_endloss_t19.push({
              value:0
            })
            }

             this.sum_ttl_marker_t20_h = 0
       this.sum_width_end_t20_h = 0
       this.result_sum_width_end_t20_h = 0
       this.sum_end_loss_length_yd_t20 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A20"){
           
           this.sum_width_end_first_t20_h = 0
            this.sum_width_end_first_t20_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t20_h = Number(this.sum_width_end_t20_h) + Number(this.sum_width_end_first_t20_h)
            this.sum_ttl_marker_t20_h = Number(this.sum_ttl_marker_t20_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t20 = Number(this.sum_end_loss_length_yd_t20) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t20_h = this.sum_width_end_t20_h / this.sum_ttl_marker_t20_h
            if(isNaN(this.result_sum_width_end_t20_h)==false){
            this.row_result_t20.push({
              value:this.result_sum_width_end_t20_h
            }) 
            }else{
              this.row_result_t20.push({
              value:0
            })
            }

        
         this.result_end_loss_yd_t20 = this.sum_end_loss_length_yd_t20 / this.sum_ttl_marker_t20_h*100
            if(isNaN(this.result_end_loss_yd_t20)==false){
            this.row_result_endloss_t20.push({
              value:this.result_end_loss_yd_t20
            })
            }else{
              this.row_result_endloss_t20.push({
              value:0
            })
            }


            this.sum_ttl_marker_tM1_h = 0
       this.sum_width_end_tM1_h = 0
       this.result_sum_width_end_tM1_h = 0
       this.sum_end_loss_length_yd_tM1 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M1"){
           
           this.sum_width_end_first_tM1_h = 0
            this.sum_width_end_first_tM1_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM1_h = Number(this.sum_width_end_tM1_h) + Number(this.sum_width_end_first_tM1_h)
            this.sum_ttl_marker_tM1_h = Number(this.sum_ttl_marker_tM1_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM1 = Number(this.sum_end_loss_length_yd_tM1) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM1_h = this.sum_width_end_tM1_h / this.sum_ttl_marker_tM1_h
            if(isNaN(this.result_sum_width_end_tM1_h)==false){
            this.row_result_tM1.push({
              value:this.result_sum_width_end_tM1_h
            }) 
            }else{
              this.row_result_tM1.push({
              value:0
            })
            }

        this.result_end_loss_yd_tM1 = this.sum_end_loss_length_yd_tM1 / this.sum_ttl_marker_tM1_h*100
            if(isNaN(this.result_end_loss_yd_tM1)==false){
            this.row_result_endloss_tM1.push({
              value:this.result_end_loss_yd_tM1
            })
            }else{
              this.row_result_endloss_tM1.push({
              value:0
            })
            }

            this.sum_ttl_marker_tM2_h = 0
       this.sum_width_end_tM2_h = 0
       this.result_sum_width_end_tM2_h = 0
       this.sum_end_loss_length_yd_tM2 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M2"){
           
           this.sum_width_end_first_tM2_h = 0
            this.sum_width_end_first_tM2_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM2_h = Number(this.sum_width_end_tM2_h) + Number(this.sum_width_end_first_tM2_h)
            this.sum_ttl_marker_tM2_h = Number(this.sum_ttl_marker_tM2_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM2 = Number(this.sum_end_loss_length_yd_tM2) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM2_h = this.sum_width_end_tM2_h / this.sum_ttl_marker_tM2_h
            if(isNaN(this.result_sum_width_end_tM2_h)==false){
            this.row_result_tM2.push({
              value:this.result_sum_width_end_tM2_h
            }) 
            }else{
              this.row_result_tM2.push({
              value:0
            })
            }

          this.result_end_loss_yd_tM2 = this.sum_end_loss_length_yd_tM2 / this.sum_ttl_marker_tM2_h*100
            if(isNaN(this.result_end_loss_yd_tM2)==false){
            this.row_result_endloss_tM2.push({
              value:this.result_end_loss_yd_tM2
            })
            }else{
              this.row_result_endloss_tM2.push({
              value:0
            })
            }
             this.sum_ttl_marker_tM3_h = 0
       this.sum_width_end_tM3_h = 0
       this.result_sum_width_end_tM3_h = 0
       this.sum_end_loss_length_yd_tM3 = 0
           
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
          
         
          if(this.rowexport_1_h[ax].gm_table == "M3"){
      
        
           this.sum_width_end_first_tM3_h = 0
     
            this.sum_width_end_first_tM3_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
         
            this.sum_width_end_tM3_h = Number(this.sum_width_end_tM3_h) + Number(this.sum_width_end_first_tM3_h)
            this.sum_ttl_marker_tM3_h = Number(this.sum_ttl_marker_tM3_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM3 = Number(this.sum_end_loss_length_yd_tM3) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM3_h = this.sum_width_end_tM3_h / this.sum_ttl_marker_tM3_h
     
            if(isNaN(this.result_sum_width_end_tM3_h)==false){
            this.row_result_tM3.push({
              value:this.result_sum_width_end_tM3_h
            }) 
            }else{
              this.row_result_tM3.push({
              value:0
            })
            }

        this.result_end_loss_yd_tM3 = this.sum_end_loss_length_yd_tM3 / this.sum_ttl_marker_tM3_h*100
            if(isNaN(this.result_end_loss_yd_tM3)==false){
            this.row_result_endloss_tM3.push({
              value:this.result_end_loss_yd_tM3
            })
            }else{
              this.row_result_endloss_tM3.push({
              value:0
            })
            }

            this.sum_ttl_marker_tM4_h = 0
       this.sum_width_end_tM4_h = 0
       this.result_sum_width_end_tM4_h = 0
       this.sum_end_loss_length_yd_tM4 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M4"){
           
           this.sum_width_end_first_tM4_h = 0
            this.sum_width_end_first_tM4_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM4_h = Number(this.sum_width_end_tM4_h) + Number(this.sum_width_end_first_tM4_h)
            this.sum_ttl_marker_tM4_h = Number(this.sum_ttl_marker_tM4_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM4 = Number(this.sum_end_loss_length_yd_tM4) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM4_h = this.sum_width_end_tM4_h / this.sum_ttl_marker_tM4_h
            if(isNaN(this.result_sum_width_end_tM4_h)==false){
            this.row_result_tM4.push({
              value:this.result_sum_width_end_tM4_h
            }) 
            }else{
              this.row_result_tM4.push({
              value:0
            })
            }

        this.result_end_loss_yd_tM4 = this.sum_end_loss_length_yd_tM4 / this.sum_ttl_marker_tM4_h*100
            if(isNaN(this.result_end_loss_yd_tM4)==false){
            this.row_result_endloss_tM4.push({
              value:this.result_end_loss_yd_tM4
            })
            }else{
              this.row_result_endloss_tM4.push({
              value:0
            })
            }

            this.sum_ttl_marker_tM5_h = 0
       this.sum_width_end_tM5_h = 0
       this.result_sum_width_end_tM5_h = 0
       this.sum_end_loss_length_yd_tM5 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M5"){
           
           this.sum_width_end_first_tM5_h = 0
            this.sum_width_end_first_tM5_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM5_h = Number(this.sum_width_end_tM5_h) + Number(this.sum_width_end_first_tM5_h)
            this.sum_ttl_marker_tM5_h = Number(this.sum_ttl_marker_tM5_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM5 = Number(this.sum_end_loss_length_yd_tM5) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM5_h = this.sum_width_end_tM5_h / this.sum_ttl_marker_tM5_h
            if(isNaN(this.result_sum_width_end_tM5_h)==false){
            this.row_result_tM5.push({
              value:this.result_sum_width_end_tM5_h
            }) 
            }else{
              this.row_result_tM5.push({
              value:0
            })
            }

          this.result_end_loss_yd_tM5 = this.sum_end_loss_length_yd_tM5 / this.sum_ttl_marker_tM5_h*100
            if(isNaN(this.result_end_loss_yd_tM5)==false){
            this.row_result_endloss_tM5.push({
              value:this.result_end_loss_yd_tM5
            })
            }else{
              this.row_result_endloss_tM5.push({
              value:0
            })
            }

      this.sum_ttl_marker_tM6_h = 0
       this.sum_width_end_tM6_h = 0
       this.result_sum_width_end_tM6_h = 0
       this.sum_end_loss_length_yd_tM6 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M6"){
           
           this.sum_width_end_first_tM6_h = 0
            this.sum_width_end_first_tM6_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM6_h = Number(this.sum_width_end_tM6_h) + Number(this.sum_width_end_first_tM6_h)
            this.sum_ttl_marker_tM6_h = Number(this.sum_ttl_marker_tM6_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM6 = Number(this.sum_end_loss_length_yd_tM6) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM6_h = this.sum_width_end_tM6_h / this.sum_ttl_marker_tM6_h
            if(isNaN(this.result_sum_width_end_tM6_h)==false){
            this.row_result_tM6.push({
              value:this.result_sum_width_end_tM6_h
            }) 
            }else{
              this.row_result_tM6.push({
              value:0
            })
            }

        this.result_end_loss_yd_tM6 = this.sum_end_loss_length_yd_tM6 / this.sum_ttl_marker_tM6_h*100
            if(isNaN(this.result_end_loss_yd_tM6)==false){
            this.row_result_endloss_tM6.push({
              value:this.result_end_loss_yd_tM6
            })
            }else{
              this.row_result_endloss_tM6.push({
              value:0
            })
            }

            this.sum_ttl_marker_tM7_h = 0
       this.sum_width_end_tM7_h = 0
       this.result_sum_width_end_tM7_h = 0
       this.sum_end_loss_length_yd_tM7 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M7"){
           
           this.sum_width_end_first_tM7_h = 0
            this.sum_width_end_first_tM7_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM7_h = Number(this.sum_width_end_tM7_h) + Number(this.sum_width_end_first_tM7_h)
            this.sum_ttl_marker_tM7_h = Number(this.sum_ttl_marker_tM7_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM7 = Number(this.sum_end_loss_length_yd_tM7) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM7_h = this.sum_width_end_tM7_h / this.sum_ttl_marker_tM7_h
            if(isNaN(this.result_sum_width_end_tM7_h)==false){
            this.row_result_tM7.push({
              value:this.result_sum_width_end_tM7_h
            }) 
            }else{
              this.row_result_tM7.push({
              value:0
            })
            }
            
          this.result_end_loss_yd_tM7 = this.sum_end_loss_length_yd_tM7 / this.sum_ttl_marker_tM7_h *100
            if(isNaN(this.result_end_loss_yd_tM7)==false){
            this.row_result_endloss_tM7.push({
              value:this.result_end_loss_yd_tM7
            })
           
          }else{
              this.row_result_endloss_tM7.push({
              value:0
            })

           
            }

             this.sum_ttl_marker_tM8_h = 0
       this.sum_width_end_tM8_h = 0
       this.result_sum_width_end_tM8_h = 0
       this.sum_end_loss_length_yd_tM8 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M8"){
           
           this.sum_width_end_first_tM8_h = 0
            this.sum_width_end_first_tM8_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM8_h = Number(this.sum_width_end_tM8_h) + Number(this.sum_width_end_first_tM8_h)
            this.sum_ttl_marker_tM8_h = Number(this.sum_ttl_marker_tM8_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM8 = Number(this.sum_end_loss_length_yd_tM8) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM8_h = this.sum_width_end_tM8_h / this.sum_ttl_marker_tM8_h
            if(isNaN(this.result_sum_width_end_tM8_h)==false){
            this.row_result_tM8.push({
              value:this.result_sum_width_end_tM8_h
            }) 
            }else{
              this.row_result_tM8.push({
              value:0
            })
            }

          this.result_end_loss_yd_tM8 = this.sum_end_loss_length_yd_tM8 / this.sum_ttl_marker_tM8_h *100
            if(isNaN(this.result_end_loss_yd_tM8)==false){
            this.row_result_endloss_tM8.push({
              value:this.result_end_loss_yd_tM8
            })
           
          }else{
              this.row_result_endloss_tM8.push({
              value:0
            })

            }

            this.sum_ttl_marker_tM9_h = 0
       this.sum_width_end_tM9_h = 0
       this.result_sum_width_end_tM9_h = 0
       this.sum_end_loss_length_yd_tM9 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M9"){
           
           this.sum_width_end_first_tM9_h = 0
            this.sum_width_end_first_tM9_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM9_h = Number(this.sum_width_end_tM9_h) + Number(this.sum_width_end_first_tM9_h)
            this.sum_ttl_marker_tM9_h = Number(this.sum_ttl_marker_tM9_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM9 = Number(this.sum_end_loss_length_yd_tM9) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM9_h = this.sum_width_end_tM9_h / this.sum_ttl_marker_tM9_h
            if(isNaN(this.result_sum_width_end_tM9_h)==false){
            this.row_result_tM9.push({
              value:this.result_sum_width_end_tM9_h
            }) 
            }else{
              this.row_result_tM9.push({
              value:0
            })
            }
            
          this.result_end_loss_yd_tM9 = this.sum_end_loss_length_yd_tM9 / this.sum_ttl_marker_tM9_h *100
            if(isNaN(this.result_end_loss_yd_tM9)==false){
            this.row_result_endloss_tM9.push({
              value:this.result_end_loss_yd_tM9
            })
           
          }else{
              this.row_result_endloss_tM9.push({
              value:0
            })

           
            }

            this.sum_ttl_marker_tM10_h = 0
       this.sum_width_end_tM10_h = 0
       this.result_sum_width_end_tM10_h = 0
       this.sum_end_loss_length_yd_tM10 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M10"){
           
           this.sum_width_end_first_tM10_h = 0
            this.sum_width_end_first_tM10_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM10_h = Number(this.sum_width_end_tM10_h) + Number(this.sum_width_end_first_tM10_h)
            this.sum_ttl_marker_tM10_h = Number(this.sum_ttl_marker_tM10_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM10 = Number(this.sum_end_loss_length_yd_tM10) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM10_h = this.sum_width_end_tM10_h / this.sum_ttl_marker_tM10_h
            if(isNaN(this.result_sum_width_end_tM10_h)==false){
            this.row_result_tM10.push({
              value:this.result_sum_width_end_tM10_h
            }) 
            }else{
              this.row_result_tM10.push({
              value:0
            })
            }
            
          this.result_end_loss_yd_tM10 = this.sum_end_loss_length_yd_tM10 / this.sum_ttl_marker_tM10_h *100
            if(isNaN(this.result_end_loss_yd_tM10)==false){
            this.row_result_endloss_tM10.push({
              value:this.result_end_loss_yd_tM10
            })
           
          }else{
              this.row_result_endloss_tM10.push({
              value:0
            })

           
            }

          }else{
            this.rowexport_1_h=[]
              this.rowexport_1_h.push(
              {
              gm_table:"",
              so_no_doc: "",
              ttl_marker:"",
              widthloss: "",
              avg_end: "",
              endless_length_yd: "",
              org:"",
              length_yd: "",
              ttl_marker:"",
              spliceloss:"",
              totalcutout:"",
              rement:"",
            })
              this.row_result_t1.push({
              value:0
              
            })
              this.row_result_t2.push({
              value:0
            })
            this.row_result_t3.push({
              value:0
            })
            this.row_result_t4.push({
              value:0
            })
            this.row_result_t5.push({
              value:0
            })
            this.row_result_t6.push({
              value:0
            })
            this.row_result_t7.push({
              value:0
            })
            this.row_result_t8.push({
              value:0
            })
            this.row_result_t9.push({
              value:0
            })
            this.row_result_t10.push({
              value:0
            })
            this.row_result_t11.push({
              value:0
            })
            this.row_result_t12.push({
              value:0
            })
            this.row_result_tM1.push({
              value:0
            })
            this.row_result_tM2.push({
              value:0
            })
            this.row_result_tM3.push({
              value:0
            })
            this.row_result_tM4.push({
              value:0
            })
            this.row_result_tM5.push({
              value:0
            })
            this.row_result_tM6.push({
              value:0
            })
            this.row_result_endloss_t1.push({
              value:0
            })
            this.row_result_endloss_t2.push({
              value:0
            })
            this.row_result_endloss_t3.push({
              value:0
            })
            this.row_result_endloss_t4.push({
              value:0
            })
            this.row_result_endloss_t5.push({
              value:0
            })
            this.row_result_endloss_t6.push({
              value:0
            })
            this.row_result_endloss_t7.push({
              value:0
            })
            this.row_result_endloss_t8.push({
              value:0
            })
            this.row_result_endloss_t9.push({
              value:0
            })
            this.row_result_endloss_t10.push({
              value:0
            })
            this.row_result_endloss_t11.push({
              value:0
            })
            this.row_result_endloss_t12.push({
              value:0
            })
            this.row_result_endloss_tM1.push({
              value:0
            })
            this.row_result_endloss_tM2.push({
              value:0
            })
            this.row_result_endloss_tM3.push({
              value:0
            })
            this.row_result_endloss_tM4.push({
              value:0
            })
            this.row_result_endloss_tM5.push({
              value:0
            })
            this.row_result_endloss_tM6.push({
              value:0
            })
           
          }
          

            //actual
            this.sum_ttl_marker_actual  = 0
            this.sum_endloss_actual = 0
            this.sum_splice_actual = 0
            this.sum_cut_actual = 0
            this.sum_rem_actual = 0

            this.last_result_actual_end = 0
            this.last_result_actual_splice = 0
            this.last_result_actual_cut = 0
            this.last_result_actual_rem = 0
          
            for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
              this.sum_ttl_marker_actual  =  Number(this.sum_ttl_marker_actual) + Number(this.rowexport_1_h[ax].ttl_marker)
              this.sum_endloss_actual =  Number(this.sum_endloss_actual) + Number(this.rowexport_1_h[ax].endless_length_yd)
              this.sum_splice_actual =  Number(this.sum_splice_actual) + Number(this.rowexport_1_h[ax].spliceloss)
              this.sum_cut_actual = Number(this.sum_cut_actual) + Number(this.rowexport_1_h[ax].totalcutout)
              this.sum_rem_actual = Number(this.sum_rem_actual) + Number(this.rowexport_1_h[ax].rement)
            }
            if(isNaN(this.sum_endloss_actual)==true){
              this.sum_endloss_actual = 0
            }
     
            this.last_result_actual_end = this.sum_endloss_actual / this.sum_ttl_marker_actual *100
            
             if(this.last_result_actual_end > 0){
              this.result_actual_end.push({
                value:this.last_result_actual_end
              })
            }else{
              this.result_actual_end.push({
                value:""
              })
            }
           
            
            if(isNaN(this.sum_splice_actual)==true){
              this.sum_splice_actual = 0
            }
            this.last_result_actual_splice = this.sum_splice_actual / this.sum_ttl_marker_actual *100
            
             
             if(this.last_result_actual_splice > 0){
              this.result_actual_splice.push({
                value:this.last_result_actual_splice
              })
            }else{
              this.result_actual_splice.push({
                value:""
              })
            }


            if(isNaN(this.sum_cut_actual)==true){
              this.sum_cut_actual = 0
            }
            this.last_result_actual_cut = this.sum_cut_actual / this.sum_ttl_marker_actual *100
          
             if(this.last_result_actual_cut > 0){
              this.result_actual_cut.push({
                value:this.last_result_actual_cut
              })
            }else{
              this.result_actual_cut.push({
                value:""
              })
            }

            if(isNaN(this.sum_rem_actual)==true){
              this.sum_rem_actual = 0
            }
            this.last_result_actual_rem = this.sum_rem_actual / this.sum_ttl_marker_actual *100
          
             if(this.last_result_actual_rem > 0){
              this.result_actual_rem.push({
                value:this.last_result_actual_rem
              })
            }else{
              this.result_actual_rem.push({
                value:""
              })
            }

          this.all_ttl_marker = 0
           this.sum_width_ttl_maker = 0
           this.sum_result_width_ttl_marker = 0
           for(var ax =0; ax<this.rowexport_1_h.length; ax++){
              this.all_ttl_marker = Number(this.all_ttl_marker) + Number(this.rowexport_1_h[ax].ttl_marker)
              this.sum_width_ttl_maker = this.rowexport_1_h[ax].ttl_marker * this.rowexport_1_h[ax].widthloss
              this.sum_result_width_ttl_marker = Number(this.sum_result_width_ttl_marker) + Number(this.sum_width_ttl_maker)
           } 
         
           this.last_result_width_ttl_marker =  this.sum_result_width_ttl_marker / this.all_ttl_marker
           
          if(this.last_result_width_ttl_marker > 0 || isNaN(this.last_result_width_ttl_marker)==false){
           this.actual_width_loss.push({
             value:this.last_result_width_ttl_marker
           })
          }else{
            this.actual_width_loss.push({
             value:0
           })
          }
 
       })
        }
       
      }else{
        this.fix_week = 52
        const params4 = new FormData();
            
       
       for(var ax = 0; ax<this.date_use_data2.length; ax++){
         
     
         params4.append("start",this.date_use_data2[ax].convert_axios1)
         params4.append("end",this.date_use_data2[ax].convert_axios2)
         params4.append("org",this.org)
        await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data_weekly_analy',
        data:params4,
       }).then((resp)=>{
         
          
          if (resp.data.data.length > 0) {
           this.rowexport_1_h=[]
          resp.data.data.forEach((e) => {
            this.rowexport_1_h.push(
              {
              mu_date:e.MU_DATE,
              gm_table: e.GM_TABLE,
              so_no_doc: e.SO_NO_DOC,
              ttl_marker: e.TTL_MARKER_YD,
              widthloss: e.WITH_LOSS,
              avg_end: e.AVG_END_LOSS,
              endless_length_yd: e.AVG_END_LOSS_YD,
              spliceloss:e.SPLICE_LOSS_YD,
              totalcutout:e.TOTAL_CUTOUT_YD,
              rement:e.REMENT_LOSS_YD,
              org:e.ORG
            })
            
          });
       
     
       
       this.sum_ttl_marker_t1_h = 0
       this.sum_width_end_t1_h = 0
       this.result_sum_width_end_t1_h = 0
       this.sum_end_loss_length_yd_t1 =0
       
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A1"){
           
           this.sum_width_end_first_t1_h = 0
            
            this.sum_width_end_first_t1_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t1_h = Number(this.sum_width_end_t1_h) + Number(this.sum_width_end_first_t1_h)
            this.sum_ttl_marker_t1_h = Number(this.sum_ttl_marker_t1_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t1 = Number(this.sum_end_loss_length_yd_t1) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
        this.result_sum_width_end_t1_h = this.sum_width_end_t1_h / this.sum_ttl_marker_t1_h
            if(isNaN(this.result_sum_width_end_t1_h)==false){
            this.row_result_t1.push({
              value:this.result_sum_width_end_t1_h
            }) 
            }else{
              this.row_result_t1.push({
              value:0
            })
            }
          
        this.result_end_loss_yd_t1 = this.sum_end_loss_length_yd_t1 / this.sum_ttl_marker_t1_h *100
            if(isNaN(this.result_end_loss_yd_t1)==false){
            this.row_result_endloss_t1.push({
              value:this.result_end_loss_yd_t1
            })
            }else{
              this.row_result_endloss_t1.push({
              value:0
            })
            }
            
            
      this.sum_ttl_marker_t2_h = 0
       this.sum_width_end_t2_h = 0
       this.result_sum_width_end_t2_h = 0
       this.sum_end_loss_length_yd_t2 = 0
       for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A2"){
           
           this.sum_width_end_first_t2_h = 0
            
            this.sum_width_end_first_t2_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t2_h = Number(this.sum_width_end_t2_h) + Number(this.sum_width_end_first_t2_h)
            this.sum_ttl_marker_t2_h = Number(this.sum_ttl_marker_t2_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t2 = Number(this.sum_end_loss_length_yd_t2) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
        this.result_sum_width_end_t2_h = this.sum_width_end_t2_h / this.sum_ttl_marker_t2_h
            if(isNaN(this.result_sum_width_end_t2_h)==false){
            this.row_result_t2.push({
              value:this.result_sum_width_end_t2_h
            }) 
            }else{
              this.row_result_t2.push({
              value:0
            })
            }

        this.result_end_loss_yd_t2 = this.sum_end_loss_length_yd_t2 / this.sum_ttl_marker_t2_h *100
            if(isNaN(this.result_end_loss_yd_t2)==false){
            this.row_result_endloss_t2.push({
              value:this.result_end_loss_yd_t2
            })
            }else{
              this.row_result_endloss_t2.push({
              value:0
            })
            }
       
        

       this.sum_ttl_marker_t3_h = 0
       this.sum_width_end_t3_h = 0
       this.result_sum_width_end_t3_h = 0
       this.sum_end_loss_length_yd_t3 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A3"){
           
           this.sum_width_end_first_t3_h = 0
            this.sum_width_end_first_t3_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t3_h = Number(this.sum_width_end_t3_h) + Number(this.sum_width_end_first_t3_h)
            this.sum_ttl_marker_t3_h = Number(this.sum_ttl_marker_t3_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t3 = Number(this.sum_end_loss_length_yd_t3) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t3_h = this.sum_width_end_t3_h / this.sum_ttl_marker_t3_h
            if(isNaN(this.result_sum_width_end_t3_h)==false){
            this.row_result_t3.push({
              value:this.result_sum_width_end_t3_h
            }) 
            }else{
              this.row_result_t3.push({
              value:0
            })
            }
        
        this.result_end_loss_yd_t3 = this.sum_end_loss_length_yd_t3 / this.sum_ttl_marker_t3_h *100
            if(isNaN(this.result_end_loss_yd_t3)==false){
            this.row_result_endloss_t3.push({
              value:this.result_end_loss_yd_t3
            })
            }else{
              this.row_result_endloss_t3.push({
              value:0
            })
            }

        
        this.sum_ttl_marker_t4_h = 0
       this.sum_width_end_t4_h = 0
       this.result_sum_width_end_t4_h = 0
       this.sum_end_loss_length_yd_t4 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A4"){
           
           this.sum_width_end_first_t4_h = 0
            this.sum_width_end_first_t4_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t4_h = Number(this.sum_width_end_t4_h) + Number(this.sum_width_end_first_t4_h)
            this.sum_ttl_marker_t4_h = Number(this.sum_ttl_marker_t4_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t4 = Number(this.sum_end_loss_length_yd_t4) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t4_h = this.sum_width_end_t4_h / this.sum_ttl_marker_t4_h
            if(isNaN(this.result_sum_width_end_t4_h)==false){
            this.row_result_t4.push({
              value:this.result_sum_width_end_t4_h
            }) 
            }else{
              this.row_result_t4.push({
              value:0
            })
            }

         this.result_end_loss_yd_t4 = this.sum_end_loss_length_yd_t4 / this.sum_ttl_marker_t4_h *100
            if(isNaN(this.result_end_loss_yd_t4)==false){
            this.row_result_endloss_t4.push({
              value:this.result_end_loss_yd_t4
            })
            }else{
              this.row_result_endloss_t4.push({
              value:0
            })
            }

       this.sum_ttl_marker_t5_h = 0
       this.sum_width_end_t5_h = 0
       this.result_sum_width_end_t5_h = 0
       this.sum_end_loss_length_yd_t5 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A5"){
           
           this.sum_width_end_first_t5_h = 0
            this.sum_width_end_first_t5_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t5_h = Number(this.sum_width_end_t5_h) + Number(this.sum_width_end_first_t5_h)
            this.sum_ttl_marker_t5_h = Number(this.sum_ttl_marker_t5_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t5 = Number(this.sum_end_loss_length_yd_t5) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t5_h = this.sum_width_end_t5_h / this.sum_ttl_marker_t5_h
            if(isNaN(this.result_sum_width_end_t5_h)==false){
            this.row_result_t5.push({
              value:this.result_sum_width_end_t5_h
            }) 
            }else{
              this.row_result_t5.push({
              value:0
            })
            }

        this.result_end_loss_yd_t5 = this.sum_end_loss_length_yd_t5 / this.sum_ttl_marker_t5_h *100
            if(isNaN(this.result_end_loss_yd_t5)==false){
            this.row_result_endloss_t5.push({
              value:this.result_end_loss_yd_t5
            })
            }else{
              this.row_result_endloss_t5.push({
              value:0
            })
            }

             this.sum_ttl_marker_t6_h = 0
       this.sum_width_end_t6_h = 0
       this.result_sum_width_end_t6_h = 0
       this.sum_end_loss_length_yd_t6 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A6"){
           
           this.sum_width_end_first_t6_h = 0
            this.sum_width_end_first_t6_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t6_h = Number(this.sum_width_end_t6_h) + Number(this.sum_width_end_first_t6_h)
            this.sum_ttl_marker_t6_h = Number(this.sum_ttl_marker_t6_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t6 = Number(this.sum_end_loss_length_yd_t6) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t6_h = this.sum_width_end_t6_h / this.sum_ttl_marker_t6_h
            if(isNaN(this.result_sum_width_end_t6_h)==false){
            this.row_result_t6.push({
              value:this.result_sum_width_end_t6_h
            }) 
            }else{
              this.row_result_t6.push({
              value:0
            })
            }

        this.result_end_loss_yd_t6 = this.sum_end_loss_length_yd_t6 / this.sum_ttl_marker_t6_h *100
            if(isNaN(this.result_end_loss_yd_t6)==false){
            this.row_result_endloss_t6.push({
              value:this.result_end_loss_yd_t6
            })
            }else{
              this.row_result_endloss_t6.push({
              value:0
            })
            }

       this.sum_ttl_marker_t7_h = 0
       this.sum_width_end_t7_h = 0
       this.result_sum_width_end_t7_h = 0
       this.sum_end_loss_length_yd_t7 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A7"){
           
           this.sum_width_end_first_t7_h = 0
            this.sum_width_end_first_t7_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t7_h = Number(this.sum_width_end_t7_h) + Number(this.sum_width_end_first_t7_h)
            this.sum_ttl_marker_t7_h = Number(this.sum_ttl_marker_t7_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t7 = Number(this.sum_end_loss_length_yd_t7) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t7_h = this.sum_width_end_t7_h / this.sum_ttl_marker_t7_h
            if(isNaN(this.result_sum_width_end_t7_h)==false){
            this.row_result_t7.push({
              value:this.result_sum_width_end_t7_h
            }) 
            }else{
              this.row_result_t7.push({
              value:0
            })
            }

        this.result_end_loss_yd_t7 = this.sum_end_loss_length_yd_t7 / this.sum_ttl_marker_t7_h *100
            if(isNaN(this.result_end_loss_yd_t7)==false){
            this.row_result_endloss_t7.push({
              value:this.result_end_loss_yd_t7
            })
            }else{
              this.row_result_endloss_t7.push({
              value:0
            })
            }
            
            this.sum_ttl_marker_t8_h = 0
       this.sum_width_end_t8_h = 0
       this.result_sum_width_end_t8_h = 0
       this.sum_end_loss_length_yd_t8 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A8"){
           
           this.sum_width_end_first_t8_h = 0
            this.sum_width_end_first_t8_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t8_h = Number(this.sum_width_end_t8_h) + Number(this.sum_width_end_first_t8_h)
            this.sum_ttl_marker_t8_h = Number(this.sum_ttl_marker_t8_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t8 = Number(this.sum_end_loss_length_yd_t8) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t8_h = this.sum_width_end_t8_h / this.sum_ttl_marker_t8_h
            if(isNaN(this.result_sum_width_end_t8_h)==false){
            this.row_result_t8.push({
              value:this.result_sum_width_end_t8_h
            }) 
            }else{
              this.row_result_t8.push({
              value:0
            })
            }

        this.result_end_loss_yd_t8 = this.sum_end_loss_length_yd_t8 / this.sum_ttl_marker_t8_h *100
            if(isNaN(this.result_end_loss_yd_t8)==false){
            this.row_result_endloss_t8.push({
              value:this.result_end_loss_yd_t8
            })
            }else{
              this.row_result_endloss_t8.push({
              value:0
            })
            }

             this.sum_ttl_marker_t9_h = 0
       this.sum_width_end_t9_h = 0
       this.result_sum_width_end_t9_h = 0
       this.sum_end_loss_length_yd_t9 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A9"){
           
           this.sum_width_end_first_t9_h = 0
            this.sum_width_end_first_t9_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t9_h = Number(this.sum_width_end_t9_h) + Number(this.sum_width_end_first_t9_h)
            this.sum_ttl_marker_t9_h = Number(this.sum_ttl_marker_t9_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t9 = Number(this.sum_end_loss_length_yd_t9) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t9_h = this.sum_width_end_t9_h / this.sum_ttl_marker_t9_h
            if(isNaN(this.result_sum_width_end_t9_h)==false){
            this.row_result_t9.push({
              value:this.result_sum_width_end_t9_h
            }) 
            }else{
              this.row_result_t9.push({
              value:0
            })
            }

        this.result_end_loss_yd_t9 = this.sum_end_loss_length_yd_t9 / this.sum_ttl_marker_t9_h *100
            if(isNaN(this.result_end_loss_yd_t9)==false){
            this.row_result_endloss_t9.push({
              value:this.result_end_loss_yd_t9
            })
            }else{
              this.row_result_endloss_t9.push({
              value:0
            })
            }

            this.sum_ttl_marker_t10_h = 0
       this.sum_width_end_t10_h = 0
       this.result_sum_width_end_t10_h = 0
       this.sum_end_loss_length_yd_t10 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A10"){
           
           this.sum_width_end_first_t10_h = 0
            this.sum_width_end_first_t10_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t10_h = Number(this.sum_width_end_t10_h) + Number(this.sum_width_end_first_t10_h)
            this.sum_ttl_marker_t10_h = Number(this.sum_ttl_marker_t10_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t10 = Number(this.sum_end_loss_length_yd_t10) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t10_h = this.sum_width_end_t10_h / this.sum_ttl_marker_t10_h
            if(isNaN(this.result_sum_width_end_t10_h)==false){
            this.row_result_t10.push({
              value:this.result_sum_width_end_t10_h
            }) 
            }else{
              this.row_result_t10.push({
              value:0
            })
            }

          this.result_end_loss_yd_t10 = this.sum_end_loss_length_yd_t10 / this.sum_ttl_marker_t10_h *100
            if(isNaN(this.result_end_loss_yd_t10)==false){
            this.row_result_endloss_t10.push({
              value:this.result_end_loss_yd_t10
            })
            }else{
              this.row_result_endloss_t10.push({
              value:0
            })
            }

            this.sum_ttl_marker_t11_h = 0
       this.sum_width_end_t11_h = 0
       this.result_sum_width_end_t11_h = 0
       this.sum_end_loss_length_yd_t11 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A11"){
           
           this.sum_width_end_first_t11_h = 0
            this.sum_width_end_first_t11_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t11_h = Number(this.sum_width_end_t11_h) + Number(this.sum_width_end_first_t11_h)
            this.sum_ttl_marker_t11_h = Number(this.sum_ttl_marker_t11_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t11 = Number(this.sum_end_loss_length_yd_t11) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t11_h = this.sum_width_end_t11_h / this.sum_ttl_marker_t11_h
            if(isNaN(this.result_sum_width_end_t11_h)==false){
            this.row_result_t11.push({
              value:this.result_sum_width_end_t11_h
            }) 
            }else{
              this.row_result_t11.push({
              value:0
            })
            }
            
        this.result_end_loss_yd_t11 = this.sum_end_loss_length_yd_t11 / this.sum_ttl_marker_t11_h *100
            if(isNaN(this.result_end_loss_yd_t11)==false){
            this.row_result_endloss_t11.push({
              value:this.result_end_loss_yd_t11
            })
            }else{
              this.row_result_endloss_t11.push({
              value:0
            })
            }

            this.sum_ttl_marker_t12_h = 0
       this.sum_width_end_t12_h = 0
       this.result_sum_width_end_t12_h = 0
       this.sum_end_loss_length_yd_t12 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A12"){
           
           this.sum_width_end_first_t12_h = 0
            this.sum_width_end_first_t12_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t12_h = Number(this.sum_width_end_t12_h) + Number(this.sum_width_end_first_t12_h)
            this.sum_ttl_marker_t12_h = Number(this.sum_ttl_marker_t12_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t12 = Number(this.sum_end_loss_length_yd_t12) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
       
         this.result_sum_width_end_t12_h = this.sum_width_end_t12_h / this.sum_ttl_marker_t12_h
         
            if(isNaN(this.result_sum_width_end_t12_h)==false){
            this.row_result_t12.push({
              value:this.result_sum_width_end_t12_h
            }) 
            }else{
              this.row_result_t12.push({
              value:0
            })
            }

        this.result_end_loss_yd_t12 = this.sum_end_loss_length_yd_t12 / this.sum_ttl_marker_t12_h *100
            if(isNaN(this.result_end_loss_yd_t12)==false){
            this.row_result_endloss_t12.push({
              value:this.result_end_loss_yd_t12
            })
            }else{
              this.row_result_endloss_t12.push({
              value:0
            })
            }

          this.sum_ttl_marker_t13_h = 0
       this.sum_width_end_t13_h = 0
       this.result_sum_width_end_t13_h = 0
       this.sum_end_loss_length_yd_t13 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A13"){
           
           this.sum_width_end_first_t13_h = 0
            this.sum_width_end_first_t13_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t13_h = Number(this.sum_width_end_t13_h) + Number(this.sum_width_end_first_t13_h)
            this.sum_ttl_marker_t13_h = Number(this.sum_ttl_marker_t13_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t13 = Number(this.sum_end_loss_length_yd_t13) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
       
         this.result_sum_width_end_t13_h = this.sum_width_end_t13_h / this.sum_ttl_marker_t13_h
            if(isNaN(this.result_sum_width_end_t13_h)==false){
            this.row_result_t13.push({
              value:this.result_sum_width_end_t13_h
            }) 
            }else{
              this.row_result_t13.push({
              value:0
            })
            }

        this.result_end_loss_yd_t13 = this.sum_end_loss_length_yd_t13 / this.sum_ttl_marker_t13_h *100
            if(isNaN(this.result_end_loss_yd_t13)==false){
            this.row_result_endloss_t13.push({
              value:this.result_end_loss_yd_t13
            })
            }else{
              this.row_result_endloss_t13.push({
              value:0
            })
            }

            this.sum_ttl_marker_t14_h = 0
       this.sum_width_end_t14_h = 0
       this.result_sum_width_end_t14_h = 0
       this.sum_end_loss_length_yd_t14 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A14"){
           
           this.sum_width_end_first_t14_h = 0
            this.sum_width_end_first_t14_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t14_h = Number(this.sum_width_end_t14_h) + Number(this.sum_width_end_first_t14_h)
            this.sum_ttl_marker_t14_h = Number(this.sum_ttl_marker_t14_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t14 = Number(this.sum_end_loss_length_yd_t14) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t14_h = this.sum_width_end_t14_h / this.sum_ttl_marker_t14_h
            if(isNaN(this.result_sum_width_end_t14_h)==false){
            this.row_result_t14.push({
              value:this.result_sum_width_end_t14_h
            }) 
            }else{
              this.row_result_t14.push({
              value:0
            })
            }

        
         this.result_end_loss_yd_t14 = this.sum_end_loss_length_yd_t14 / this.sum_ttl_marker_t14_h*100
            if(isNaN(this.result_end_loss_yd_t14)==false){
            this.row_result_endloss_t14.push({
              value:this.result_end_loss_yd_t14
            })
            }else{
              this.row_result_endloss_t14.push({
              value:0
            })
            }

            this.sum_ttl_marker_t15_h = 0
       this.sum_width_end_t15_h = 0
       this.result_sum_width_end_t15_h = 0
       this.sum_end_loss_length_yd_t15 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A15"){
           
           this.sum_width_end_first_t15_h = 0
            this.sum_width_end_first_t15_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t15_h = Number(this.sum_width_end_t15_h) + Number(this.sum_width_end_first_t15_h)
            this.sum_ttl_marker_t15_h = Number(this.sum_ttl_marker_t15_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t15 = Number(this.sum_end_loss_length_yd_t15) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t15_h = this.sum_width_end_t15_h / this.sum_ttl_marker_t15_h
            if(isNaN(this.result_sum_width_end_t15_h)==false){
            this.row_result_t15.push({
              value:this.result_sum_width_end_t15_h
            }) 
            }else{
              this.row_result_t15.push({
              value:0
            })
            }

        
         this.result_end_loss_yd_t15 = this.sum_end_loss_length_yd_t15 / this.sum_ttl_marker_t15_h*100
            if(isNaN(this.result_end_loss_yd_t15)==false){
            this.row_result_endloss_t15.push({
              value:this.result_end_loss_yd_t15
            })
            }else{
              this.row_result_endloss_t15.push({
              value:0
            })
            }

            this.sum_ttl_marker_t16_h = 0
       this.sum_width_end_t16_h = 0
       this.result_sum_width_end_t16_h = 0
       this.sum_end_loss_length_yd_t16 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A16"){
           
           this.sum_width_end_first_t16_h = 0
            this.sum_width_end_first_t16_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t16_h = Number(this.sum_width_end_t16_h) + Number(this.sum_width_end_first_t16_h)
            this.sum_ttl_marker_t16_h = Number(this.sum_ttl_marker_t16_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t16 = Number(this.sum_end_loss_length_yd_t16) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t16_h = this.sum_width_end_t16_h / this.sum_ttl_marker_t16_h
            if(isNaN(this.result_sum_width_end_t16_h)==false){
            this.row_result_t16.push({
              value:this.result_sum_width_end_t16_h
            }) 
            }else{
              this.row_result_t16.push({
              value:0
            })
            }

        
         this.result_end_loss_yd_t16 = this.sum_end_loss_length_yd_t16 / this.sum_ttl_marker_t16_h*100
            if(isNaN(this.result_end_loss_yd_t16)==false){
            this.row_result_endloss_t16.push({
              value:this.result_end_loss_yd_t16
            })
            }else{
              this.row_result_endloss_t16.push({
              value:0
            })
            }

            this.sum_ttl_marker_t17_h = 0
       this.sum_width_end_t17_h = 0
       this.result_sum_width_end_t17_h = 0
       this.sum_end_loss_length_yd_t17 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A17"){
           
           this.sum_width_end_first_t17_h = 0
            this.sum_width_end_first_t17_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t17_h = Number(this.sum_width_end_t17_h) + Number(this.sum_width_end_first_t17_h)
            this.sum_ttl_marker_t17_h = Number(this.sum_ttl_marker_t17_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t17 = Number(this.sum_end_loss_length_yd_t17) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t17_h = this.sum_width_end_t17_h / this.sum_ttl_marker_t17_h
            if(isNaN(this.result_sum_width_end_t17_h)==false){
            this.row_result_t17.push({
              value:this.result_sum_width_end_t17_h
            }) 
            }else{
              this.row_result_t17.push({
              value:0
            })
            }

        
         this.result_end_loss_yd_t17 = this.sum_end_loss_length_yd_t17 / this.sum_ttl_marker_t17_h*100
            if(isNaN(this.result_end_loss_yd_t17)==false){
            this.row_result_endloss_t17.push({
              value:this.result_end_loss_yd_t17
            })
            }else{
              this.row_result_endloss_t17.push({
              value:0
            })
            }

            this.sum_ttl_marker_t18_h = 0
       this.sum_width_end_t18_h = 0
       this.result_sum_width_end_t18_h = 0
       this.sum_end_loss_length_yd_t18 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A18"){
           
           this.sum_width_end_first_t18_h = 0
            this.sum_width_end_first_t18_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t18_h = Number(this.sum_width_end_t18_h) + Number(this.sum_width_end_first_t18_h)
            this.sum_ttl_marker_t18_h = Number(this.sum_ttl_marker_t18_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t18 = Number(this.sum_end_loss_length_yd_t18) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t18_h = this.sum_width_end_t18_h / this.sum_ttl_marker_t18_h
            if(isNaN(this.result_sum_width_end_t18_h)==false){
            this.row_result_t18.push({
              value:this.result_sum_width_end_t18_h
            }) 
            }else{
              this.row_result_t18.push({
              value:0
            })
            }

        
         this.result_end_loss_yd_t18 = this.sum_end_loss_length_yd_t18 / this.sum_ttl_marker_t18_h*100
            if(isNaN(this.result_end_loss_yd_t18)==false){
            this.row_result_endloss_t18.push({
              value:this.result_end_loss_yd_t18
            })
            }else{
              this.row_result_endloss_t18.push({
              value:0
            })
            }

            this.sum_ttl_marker_t19_h = 0
       this.sum_width_end_t19_h = 0
       this.result_sum_width_end_t19_h = 0
       this.sum_end_loss_length_yd_t19 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A19"){
           
           this.sum_width_end_first_t19_h = 0
            this.sum_width_end_first_t19_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t19_h = Number(this.sum_width_end_t19_h) + Number(this.sum_width_end_first_t19_h)
            this.sum_ttl_marker_t19_h = Number(this.sum_ttl_marker_t19_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t19 = Number(this.sum_end_loss_length_yd_t19) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t19_h = this.sum_width_end_t19_h / this.sum_ttl_marker_t19_h
            if(isNaN(this.result_sum_width_end_t19_h)==false){
            this.row_result_t19.push({
              value:this.result_sum_width_end_t19_h
            }) 
            }else{
              this.row_result_t19.push({
              value:0
            })
            }

        
         this.result_end_loss_yd_t19 = this.sum_end_loss_length_yd_t19 / this.sum_ttl_marker_t19_h*100
            if(isNaN(this.result_end_loss_yd_t19)==false){
            this.row_result_endloss_t19.push({
              value:this.result_end_loss_yd_t19
            })
            }else{
              this.row_result_endloss_t19.push({
              value:0
            })
            }

             this.sum_ttl_marker_t20_h = 0
       this.sum_width_end_t20_h = 0
       this.result_sum_width_end_t20_h = 0
       this.sum_end_loss_length_yd_t20 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "A20"){
           
           this.sum_width_end_first_t20_h = 0
            this.sum_width_end_first_t20_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_t20_h = Number(this.sum_width_end_t20_h) + Number(this.sum_width_end_first_t20_h)
            this.sum_ttl_marker_t20_h = Number(this.sum_ttl_marker_t20_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_t20 = Number(this.sum_end_loss_length_yd_t20) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_t20_h = this.sum_width_end_t20_h / this.sum_ttl_marker_t20_h
            if(isNaN(this.result_sum_width_end_t20_h)==false){
            this.row_result_t20.push({
              value:this.result_sum_width_end_t20_h
            }) 
            }else{
              this.row_result_t20.push({
              value:0
            })
            }

        
         this.result_end_loss_yd_t20 = this.sum_end_loss_length_yd_t20 / this.sum_ttl_marker_t20_h*100
            if(isNaN(this.result_end_loss_yd_t20)==false){
            this.row_result_endloss_t20.push({
              value:this.result_end_loss_yd_t20
            })
            }else{
              this.row_result_endloss_t20.push({
              value:0
            })
            }

            this.sum_ttl_marker_tM1_h = 0
       this.sum_width_end_tM1_h = 0
       this.result_sum_width_end_tM1_h = 0
       this.sum_end_loss_length_yd_tM1 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M1"){
           
           this.sum_width_end_first_tM1_h = 0
            this.sum_width_end_first_tM1_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM1_h = Number(this.sum_width_end_tM1_h) + Number(this.sum_width_end_first_tM1_h)
            this.sum_ttl_marker_tM1_h = Number(this.sum_ttl_marker_tM1_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM1 = Number(this.sum_end_loss_length_yd_tM1) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM1_h = this.sum_width_end_tM1_h / this.sum_ttl_marker_tM1_h
            if(isNaN(this.result_sum_width_end_tM1_h)==false){
            this.row_result_tM1.push({
              value:this.result_sum_width_end_tM1_h
            }) 
            }else{
              this.row_result_tM1.push({
              value:0
            })
            }

        this.result_end_loss_yd_tM1 = this.sum_end_loss_length_yd_tM1 / this.sum_ttl_marker_tM1_h *100
            if(isNaN(this.result_end_loss_yd_tM1)==false){
            this.row_result_endloss_tM1.push({
              value:this.result_end_loss_yd_tM1
            })
            }else{
              this.row_result_endloss_tM1.push({
              value:0
            })
            }

            this.sum_ttl_marker_tM2_h = 0
       this.sum_width_end_tM2_h = 0
       this.result_sum_width_end_tM2_h = 0
       this.sum_end_loss_length_yd_tM2 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M2"){
           
           this.sum_width_end_first_tM2_h = 0
            this.sum_width_end_first_tM2_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM2_h = Number(this.sum_width_end_tM2_h) + Number(this.sum_width_end_first_tM2_h)
            this.sum_ttl_marker_tM2_h = Number(this.sum_ttl_marker_tM2_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM2 = Number(this.sum_end_loss_length_yd_tM2) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM2_h = this.sum_width_end_tM2_h / this.sum_ttl_marker_tM2_h
            if(isNaN(this.result_sum_width_end_tM2_h)==false){
            this.row_result_tM2.push({
              value:this.result_sum_width_end_tM2_h
            }) 
            }else{
              this.row_result_tM2.push({
              value:0
            })
            }

         this.result_end_loss_yd_tM2 = this.sum_end_loss_length_yd_tM2 / this.sum_ttl_marker_tM2_h *100
            if(isNaN(this.result_end_loss_yd_tM2)==false){
            this.row_result_endloss_tM2.push({
              value:this.result_end_loss_yd_tM2
            })
            }else{
              this.row_result_endloss_tM2.push({
              value:0
            })
            }
             this.sum_ttl_marker_tM3_h = 0
       this.sum_width_end_tM3_h = 0
       this.result_sum_width_end_tM3_h = 0
       this.sum_end_loss_length_yd_tM3 = 0
           
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
          
         
          if(this.rowexport_1_h[ax].gm_table == "M3"){
      
        
           this.sum_width_end_first_tM3_h = 0
     
            this.sum_width_end_first_tM3_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
         
            this.sum_width_end_tM3_h = Number(this.sum_width_end_tM3_h) + Number(this.sum_width_end_first_tM3_h)
            this.sum_ttl_marker_tM3_h = Number(this.sum_ttl_marker_tM3_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM3 = Number(this.sum_end_loss_length_yd_tM3) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM3_h = this.sum_width_end_tM3_h / this.sum_ttl_marker_tM3_h
     
            if(isNaN(this.result_sum_width_end_tM3_h)==false){
            this.row_result_tM3.push({
              value:this.result_sum_width_end_tM3_h
            }) 
            }else{
              this.row_result_tM3.push({
              value:0
            })
            }

          this.result_end_loss_yd_tM3 = this.sum_end_loss_length_yd_tM3 / this.sum_ttl_marker_tM3_h *100
            if(isNaN(this.result_end_loss_yd_tM3)==false){
            this.row_result_endloss_tM3.push({
              value:this.result_end_loss_yd_tM3
            })
            }else{
              this.row_result_endloss_tM3.push({
              value:0
            })
            }

        

       

            this.sum_ttl_marker_tM4_h = 0
       this.sum_width_end_tM4_h = 0
       this.result_sum_width_end_tM4_h = 0
       this.sum_end_loss_length_yd_tM4 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M4"){
           
           this.sum_width_end_first_tM4_h = 0
            this.sum_width_end_first_tM4_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM4_h = Number(this.sum_width_end_tM4_h) + Number(this.sum_width_end_first_tM4_h)
            this.sum_ttl_marker_tM4_h = Number(this.sum_ttl_marker_tM4_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM4 = Number(this.sum_end_loss_length_yd_tM4) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM4_h = this.sum_width_end_tM4_h / this.sum_ttl_marker_tM4_h
            if(isNaN(this.result_sum_width_end_tM4_h)==false){
            this.row_result_tM4.push({
              value:this.result_sum_width_end_tM4_h
            }) 
            }else{
              this.row_result_tM4.push({
              value:0
            })
            }

        this.result_end_loss_yd_tM4 = this.sum_end_loss_length_yd_tM4 / this.sum_ttl_marker_tM4_h *100
            if(isNaN(this.result_end_loss_yd_tM4)==false){
            this.row_result_endloss_tM4.push({
              value:this.result_end_loss_yd_tM4
            })
            }else{
              this.row_result_endloss_tM4.push({
              value:0
            })
            }

            this.sum_ttl_marker_tM5_h = 0
       this.sum_width_end_tM5_h = 0
       this.result_sum_width_end_tM5_h = 0
       this.sum_end_loss_length_yd_tM5 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M5"){
           
           this.sum_width_end_first_tM5_h = 0
            this.sum_width_end_first_tM5_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM5_h = Number(this.sum_width_end_tM5_h) + Number(this.sum_width_end_first_tM5_h)
            this.sum_ttl_marker_tM5_h = Number(this.sum_ttl_marker_tM5_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM5 = Number(this.sum_end_loss_length_yd_tM5) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM5_h = this.sum_width_end_tM5_h / this.sum_ttl_marker_tM5_h
            if(isNaN(this.result_sum_width_end_tM5_h)==false){
            this.row_result_tM5.push({
              value:this.result_sum_width_end_tM5_h
            }) 
            }else{
              this.row_result_tM5.push({
              value:0
            })
            }

         this.result_end_loss_yd_tM5 = this.sum_end_loss_length_yd_tM5 / this.sum_ttl_marker_tM5_h *100
            if(isNaN(this.result_end_loss_yd_tM5)==false){
            this.row_result_endloss_tM5.push({
              value:this.result_end_loss_yd_tM5
            })
            }else{
              this.row_result_endloss_tM5.push({
              value:0
            })
            }

            this.sum_ttl_marker_tM6_h = 0
       this.sum_width_end_tM6_h = 0
       this.result_sum_width_end_tM6_h = 0
       this.sum_end_loss_length_yd_tM6 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M6"){
           
           this.sum_width_end_first_tM6_h = 0
            this.sum_width_end_first_tM6_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM6_h = Number(this.sum_width_end_tM6_h) + Number(this.sum_width_end_first_tM6_h)
            this.sum_ttl_marker_tM6_h = Number(this.sum_ttl_marker_tM6_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM6 = Number(this.sum_end_loss_length_yd_tM6) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM6_h = this.sum_width_end_tM6_h / this.sum_ttl_marker_tM6_h
            if(isNaN(this.result_sum_width_end_tM6_h)==false){
            this.row_result_tM6.push({
              value:this.result_sum_width_end_tM6_h
            }) 
            }else{
              this.row_result_tM6.push({
              value:0
            })
            }
            
          this.result_end_loss_yd_tM6 = this.sum_end_loss_length_yd_tM6 / this.sum_ttl_marker_tM6_h *100
            if(isNaN(this.result_end_loss_yd_tM6)==false){
            this.row_result_endloss_tM6.push({
              value:this.result_end_loss_yd_tM6
            })
           
          }else{
              this.row_result_endloss_tM6.push({
              value:0
            })

           
            }


      this.sum_ttl_marker_tM7_h = 0
       this.sum_width_end_tM7_h = 0
       this.result_sum_width_end_tM7_h = 0
       this.sum_end_loss_length_yd_tM7 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M7"){
           
           this.sum_width_end_first_tM7_h = 0
            this.sum_width_end_first_tM7_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM7_h = Number(this.sum_width_end_tM7_h) + Number(this.sum_width_end_first_tM7_h)
            this.sum_ttl_marker_tM7_h = Number(this.sum_ttl_marker_tM7_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM7 = Number(this.sum_end_loss_length_yd_tM7) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM7_h = this.sum_width_end_tM7_h / this.sum_ttl_marker_tM7_h
            if(isNaN(this.result_sum_width_end_tM7_h)==false){
            this.row_result_tM7.push({
              value:this.result_sum_width_end_tM7_h
            }) 
            }else{
              this.row_result_tM7.push({
              value:0
            })
            }
            
          this.result_end_loss_yd_tM7 = this.sum_end_loss_length_yd_tM7 / this.sum_ttl_marker_tM7_h *100
            if(isNaN(this.result_end_loss_yd_tM7)==false){
            this.row_result_endloss_tM7.push({
              value:this.result_end_loss_yd_tM7
            })
           
          }else{
              this.row_result_endloss_tM7.push({
              value:0
            })

           
            }

             this.sum_ttl_marker_tM8_h = 0
       this.sum_width_end_tM8_h = 0
       this.result_sum_width_end_tM8_h = 0
       this.sum_end_loss_length_yd_tM8 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M8"){
           
           this.sum_width_end_first_tM8_h = 0
            this.sum_width_end_first_tM8_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM8_h = Number(this.sum_width_end_tM8_h) + Number(this.sum_width_end_first_tM8_h)
            this.sum_ttl_marker_tM8_h = Number(this.sum_ttl_marker_tM8_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM8 = Number(this.sum_end_loss_length_yd_tM8) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM8_h = this.sum_width_end_tM8_h / this.sum_ttl_marker_tM8_h
            if(isNaN(this.result_sum_width_end_tM8_h)==false){
            this.row_result_tM8.push({
              value:this.result_sum_width_end_tM8_h
            }) 
            }else{
              this.row_result_tM8.push({
              value:0
            })
            }

          this.result_end_loss_yd_tM8 = this.sum_end_loss_length_yd_tM8 / this.sum_ttl_marker_tM8_h *100
            if(isNaN(this.result_end_loss_yd_tM8)==false){
            this.row_result_endloss_tM8.push({
              value:this.result_end_loss_yd_tM8
            })
           
          }else{
              this.row_result_endloss_tM8.push({
              value:0
            })

            }

            this.sum_ttl_marker_tM9_h = 0
       this.sum_width_end_tM9_h = 0
       this.result_sum_width_end_tM9_h = 0
       this.sum_end_loss_length_yd_tM9 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M9"){
           
           this.sum_width_end_first_tM9_h = 0
            this.sum_width_end_first_tM9_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM9_h = Number(this.sum_width_end_tM9_h) + Number(this.sum_width_end_first_tM9_h)
            this.sum_ttl_marker_tM9_h = Number(this.sum_ttl_marker_tM9_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM9 = Number(this.sum_end_loss_length_yd_tM9) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM9_h = this.sum_width_end_tM9_h / this.sum_ttl_marker_tM9_h
            if(isNaN(this.result_sum_width_end_tM9_h)==false){
            this.row_result_tM9.push({
              value:this.result_sum_width_end_tM9_h
            }) 
            }else{
              this.row_result_tM9.push({
              value:0
            })
            }
            
          this.result_end_loss_yd_tM9 = this.sum_end_loss_length_yd_tM9 / this.sum_ttl_marker_tM9_h *100
            if(isNaN(this.result_end_loss_yd_tM9)==false){
            this.row_result_endloss_tM9.push({
              value:this.result_end_loss_yd_tM9
            })
           
          }else{
              this.row_result_endloss_tM9.push({
              value:0
            })

           
            }

            this.sum_ttl_marker_tM10_h = 0
       this.sum_width_end_tM10_h = 0
       this.result_sum_width_end_tM10_h = 0
       this.sum_end_loss_length_yd_tM10 = 0
       
        for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
       
     
          if(this.rowexport_1_h[ax].gm_table == "M10"){
           
           this.sum_width_end_first_tM10_h = 0
            this.sum_width_end_first_tM10_h = this.rowexport_1_h[ax].widthloss * this.rowexport_1_h[ax].ttl_marker
            this.sum_width_end_tM10_h = Number(this.sum_width_end_tM10_h) + Number(this.sum_width_end_first_tM10_h)
            this.sum_ttl_marker_tM10_h = Number(this.sum_ttl_marker_tM10_h) + Number(this.rowexport_1_h[ax].ttl_marker) //ผลรวม ttl */
            this.sum_end_loss_length_yd_tM10 = Number(this.sum_end_loss_length_yd_tM10) + Number(this.rowexport_1_h[ax].endless_length_yd)

            }
        
        }  
         this.result_sum_width_end_tM10_h = this.sum_width_end_tM10_h / this.sum_ttl_marker_tM10_h
            if(isNaN(this.result_sum_width_end_tM10_h)==false){
            this.row_result_tM10.push({
              value:this.result_sum_width_end_tM10_h
            }) 
            }else{
              this.row_result_tM10.push({
              value:0
            })
            }
            
          this.result_end_loss_yd_tM10 = this.sum_end_loss_length_yd_tM10 / this.sum_ttl_marker_tM10_h *100
            if(isNaN(this.result_end_loss_yd_tM10)==false){
            this.row_result_endloss_tM10.push({
              value:this.result_end_loss_yd_tM10
            })
           
          }else{
              this.row_result_endloss_tM10.push({
              value:0
            })

           
            }
          }else{
            this.rowexport_1_h=[]
            this.rowexport_1_h.push(
              {
              mu_date:"",
              gm_table:"",
              so_no_doc: "",
              ttl_marker: "",
              widthloss: "",
              avg_end: "",
              endless_length_yd:"",
              org:""
            })
            this.row_result_t1.push({
              value:0
            })
            this.row_result_t2.push({
              value:0
            })
            this.row_result_t3.push({
              value:0
            })
            this.row_result_t4.push({
              value:0
            })
            this.row_result_t5.push({
              value:0
            })
            this.row_result_t6.push({
              value:0
            })
            this.row_result_t7.push({
              value:0
            })
            this.row_result_t8.push({
              value:0
            })

            this.row_result_t9.push({
              value:0
            })
            this.row_result_t10.push({
              value:0
            })
            this.row_result_t11.push({
              value:0
            })
            this.row_result_t12.push({
              value:0
            })
            this.row_result_t13.push({
              value:0
            })
            this.row_result_t14.push({
              value:0
            })
            this.row_result_t15.push({
              value:0
            })
            this.row_result_t16.push({
              value:0
            })
            this.row_result_t17.push({
              value:0
            })
            this.row_result_t18.push({
              value:0
            })
            this.row_result_t19.push({
              value:0
            })
            this.row_result_t20.push({
              value:0
            })
            this.row_result_tM1.push({
              value:0
            })
            this.row_result_tM2.push({
              value:0
            })
            this.row_result_tM3.push({
              value:0
            })
            this.row_result_tM4.push({
              value:0
            })
            this.row_result_tM5.push({
              value:0
            })
            this.row_result_tM6.push({
              value:0
            })
            this.row_result_tM7.push({
              value:0
            })
            this.row_result_tM8.push({
              value:0
            })
            this.row_result_tM9.push({
              value:0
            })
            this.row_result_tM10.push({
              value:0
            })
            
            this.row_result_endloss_t1.push({
              value:0
            })
            this.row_result_endloss_t2.push({
              value:0
            })
            this.row_result_endloss_t3.push({
              value:0
            })
            this.row_result_endloss_t4.push({
              value:0
            })
            this.row_result_endloss_t5.push({
              value:0
            })
            this.row_result_endloss_t6.push({
              value:0
            })
            this.row_result_endloss_t7.push({
              value:0
            })
            this.row_result_endloss_t8.push({
              value:0
            })
            this.row_result_endloss_t9.push({
              value:0
            })
            this.row_result_endloss_t10.push({
              value:0
            })
            this.row_result_endloss_t11.push({
              value:0
            })
            this.row_result_endloss_t12.push({
              value:0
            })
             this.row_result_endloss_t13.push({
              value:0
            })
             this.row_result_endloss_t14.push({
              value:0
            })
             this.row_result_endloss_t15.push({
              value:0
            })
             this.row_result_endloss_t16.push({
              value:0
            })
             this.row_result_endloss_t17.push({
              value:0
            })
             this.row_result_endloss_t18.push({
              value:0
            })
             this.row_result_endloss_t19.push({
              value:0
            })
             this.row_result_endloss_t20.push({
              value:0
            })
            this.row_result_endloss_tM1.push({
              value:0
            })
            this.row_result_endloss_tM2.push({
              value:0
            })
            this.row_result_endloss_tM3.push({
              value:0
            })
            this.row_result_endloss_tM4.push({
              value:0
            })
            this.row_result_endloss_tM5.push({
              value:0
            })
            this.row_result_endloss_tM6.push({
              value:0
            })
            this.row_result_endloss_tM7.push({
              value:0
            })
            this.row_result_endloss_tM8.push({
              value:0
            })
            this.row_result_endloss_tM9.push({
              value:0
            })
            this.row_result_endloss_tM10.push({
              value:0
            })
            

          }
           //actual
           this.last_result_actual_end = 0
            this.last_result_actual_splice = 0
            this.last_result_actual_cut = 0
            this.last_result_actual_rem = 0
         
            this.sum_ttl_marker_actual = 0
            this.sum_endloss_actual = 0
            this.sum_splice_actual = 0
            this.sum_cut_actual = 0
            this.sum_rem_actual = 0
           
            for(var ax = 0; ax<this.rowexport_1_h.length; ax++){
              this.sum_ttl_marker_actual  =  Number(this.sum_ttl_marker_actual) + Number(this.rowexport_1_h[ax].ttl_marker)
              this.sum_endloss_actual =  Number(this.sum_endloss_actual) + Number(this.rowexport_1_h[ax].endless_length_yd)
              this.sum_splice_actual =  Number(this.sum_splice_actual) + Number(this.rowexport_1_h[ax].spliceloss)
              this.sum_cut_actual = Number(this.sum_cut_actual) + Number(this.rowexport_1_h[ax].totalcutout)
              this.sum_rem_actual = Number(this.sum_rem_actual) + Number(this.rowexport_1_h[ax].rement)
            }
           



             if(isNaN(this.sum_endloss_actual)==true){
              this.sum_endloss_actual = 0
            }
           
            this.last_result_actual_end = this.sum_endloss_actual / this.sum_ttl_marker_actual *100
            
             if(this.last_result_actual_end > 0){
              this.result_actual_end.push({
                value:this.last_result_actual_end
              })
            }else{
              this.result_actual_end.push({
                value:""
              })
            }
           
            
            if(isNaN(this.sum_splice_actual)==true){
              this.sum_splice_actual = 0
            }
            this.last_result_actual_splice = this.sum_splice_actual / this.sum_ttl_marker_actual *100
     
             if(this.last_result_actual_splice > 0){
              this.result_actual_splice.push({
                value:this.last_result_actual_splice
              })
            }else{
              this.result_actual_splice.push({
                value:""
              })
            }


            if(isNaN(this.sum_cut_actual)==true){
              this.sum_cut_actual = 0
            }
            this.last_result_actual_cut = this.sum_cut_actual / this.sum_ttl_marker_actual *100
     
             if(this.last_result_actual_cut > 0){
              this.result_actual_cut.push({
                value:this.last_result_actual_cut
              })
            }else{
              this.result_actual_cut.push({
                value:""
              })
            }

            if(isNaN(this.sum_rem_actual)==true){
              this.sum_rem_actual = 0
            }
            this.last_result_actual_rem = this.sum_rem_actual / this.sum_ttl_marker_actual *100
           
             if(this.last_result_actual_rem > 0){
              this.result_actual_rem.push({
                value:this.last_result_actual_rem
              })
            }else{
              this.result_actual_rem.push({
                value:""
              })
            }

            this.all_ttl_marker = 0
           this.sum_width_ttl_maker = 0
           this.sum_result_width_ttl_marker = 0
           for(var ax =0; ax<this.rowexport_1_h.length; ax++){
              this.all_ttl_marker = Number(this.all_ttl_marker) + Number(this.rowexport_1_h[ax].ttl_marker)
              this.sum_width_ttl_maker = this.rowexport_1_h[ax].ttl_marker * this.rowexport_1_h[ax].widthloss
              this.sum_result_width_ttl_marker = Number(this.sum_result_width_ttl_marker) + Number(this.sum_width_ttl_maker)
           } 
           this.last_result_width_ttl_marker =  this.sum_result_width_ttl_marker / this.all_ttl_marker
           
          if(this.last_result_width_ttl_marker > 0 || isNaN(this.last_result_width_ttl_marker)==false){
           this.actual_width_loss.push({
             value:this.last_result_width_ttl_marker
           })
          }else{
            this.actual_width_loss.push({
             value:0
           })
          }
          

            
            
          
       })
       } 


      }



//sheet4
      ws2.getCell("A1").value = "Table"

      ws2.getCell("A2").value = "Monthly"
      ws2.mergeCells("A2:A3");

      ws2.getCell("B1").value = "T.A1"

      ws2.getCell("C1").value = "T.A2"

      ws2.getCell("D1").value = "T.A3"

      ws2.getCell("E1").value = "T.A4"

      ws2.getCell("F1").value = "T.A5"
      
      ws2.getCell("G1").value = "T.A6"

      ws2.getCell("H1").value = "T.A7"

      ws2.getCell("I1").value = "T.A8"

      ws2.getCell("J1").value = "T.A9"

      ws2.getCell("K1").value = "T.A10"
      
      ws2.getCell("L1").value = "T.A11"

      ws2.getCell("M1").value = "T.A12"

      ws2.getCell("N1").value = "T.A13"

      ws2.getCell("O1").value = "T.A14"
      
      ws2.getCell("P1").value = "T.A15"

      ws2.getCell("Q1").value = "T.A16"

      ws2.getCell("R1").value = "T.A17"

      ws2.getCell("S1").value = "T.A18"

      ws2.getCell("T1").value = "T.A19"

      ws2.getCell("U1").value = "T.A20"

      ws2.getCell("V1").value = "T.M1"

      ws2.getCell("W1").value = "T.M2"

      ws2.getCell("X1").value = "T.M3"

      ws2.getCell("Y1").value = "T.M4"

      ws2.getCell("Z1").value = "T.M5"

      ws2.getCell("AA1").value = "T.M6"

      ws2.getCell("AB1").value = "T.M7"

      ws2.getCell("AC1").value = "T.M8"

      ws2.getCell("AD1").value = "T.M9"

      ws2.getCell("AE1").value = "T.M10"

      

      ws2.getCell("B2").value = "Width Loss%"
      ws2.mergeCells("B2:AE2")

      ws2.getCell("AF1").value = "Target"

     
      ws2.getCell("AG1").value = "T.A1"

      ws2.getCell("AH1").value = "T.A2"
      
      ws2.getCell("AI1").value = "T.A3"

      ws2.getCell("AJ1").value = "T.A4"

      ws2.getCell("AK1").value = "T.A5"

      ws2.getCell("AL1").value = "T.A6"

      ws2.getCell("AM1").value = "T.A7"

      ws2.getCell("AN1").value = "T.A8"

      ws2.getCell("AO1").value = "T.A9"

      ws2.getCell("AP1").value = "T.A10"

      ws2.getCell("AQ1").value = "T.A11"

      ws2.getCell("AR1").value = "T.A12"

      ws2.getCell("AS1").value = "T.A13"

      ws2.getCell("AT1").value = "T.A14"

      ws2.getCell("AU1").value = "T.A15"

      ws2.getCell("AV1").value = "T.A16"

      ws2.getCell("AW1").value = "T.A17"

      ws2.getCell("AX1").value = "T.A18"

      ws2.getCell("AY1").value = "T.A19"

      ws2.getCell("AZ1").value = "T.A20"

      ws2.getCell("BA1").value = "T.M1"
      
      ws2.getCell("BB1").value = "T.M2"

      ws2.getCell("BC1").value = "T.M3"

      ws2.getCell("BD1").value = "T.M4"

      ws2.getCell("BE1").value = "T.M5"

      ws2.getCell("BF1").value = "T.M6"

      ws2.getCell("BG1").value = "T.M7"

      ws2.getCell("BH1").value = "T.M8"

      ws2.getCell("BI1").value = "T.M9"

      ws2.getCell("BJ1").value = "T.M10"


     ws2.getCell("AG2").value = "EndLoss%"
      ws2.mergeCells("AG2:BJ2")

      ws2.getCell("BK1").value = ""

      ws2.getCell("BL1").value = "End Loss"

      ws2.getCell("BM1").value = "Splice Loss"

      ws2.getCell("BN1").value = "Cut Out Loss"

      ws2.getCell("BO1").value = "Remnants Loss"

      ws2.getCell("BP1").value = "Total Length Loss"

      ws2.getCell("BQ1").value = "Width Loss"

      ws2.getCell("BR3").value = "Actual"
      
      ws2.mergeCells("BK3:BR3")


      ws2.getCell("BS1").value = "End Loss"

      ws2.getCell("BT1").value = "Splice Loss"

      ws2.getCell("BU1").value = "Cut Out Loss"

      ws2.getCell("BV1").value = "Remnants Loss"

      ws2.getCell("BW1").value = "Total Length Loss"


      ws2.getCell("BS3").value = "Target"
      ws2.mergeCells("BS3:BW3")

      if(this.start_h !== ""){
         for(var bz = 4; bz<this.date_use_data.length+4;  bz++){
           ws2.getCell("BS"+[bz]).value = 0.34/100
           ws2.getCell("BS"+[bz]).numFmt = "0.00%"

           ws2.getCell("BT"+[bz]).value = 0.10/100
           ws2.getCell("BT"+[bz]).numFmt = "0.00%"

           ws2.getCell("BU"+[bz]).value = 0.00/100
           ws2.getCell("BU"+[bz]).numFmt = "0.00%"

           ws2.getCell("BV"+[bz]).value = 0.08/100
           ws2.getCell("BV"+[bz]).numFmt = "0.00%"

           ws2.getCell("BW"+[bz]).value = 0.52/100
           ws2.getCell("BW"+[bz]).numFmt = "0.00%"
         }
      }else{
        for(var bz = 4; bz<this.date_use_data2.length+4;  bz++){
           ws2.getCell("BS"+[bz]).value = 0.34/100
           ws2.getCell("BS"+[bz]).numFmt = "0.00%"

           ws2.getCell("BT"+[bz]).value = 0.10/100
           ws2.getCell("BT"+[bz]).numFmt = "0.00%"

           ws2.getCell("BU"+[bz]).value = 0.00/100
           ws2.getCell("BU"+[bz]).numFmt = "0.00%"

           ws2.getCell("BV"+[bz]).value = 0.08/100
           ws2.getCell("BV"+[bz]).numFmt = "0.00%"

           ws2.getCell("BW"+[bz]).value = 0.52/100
           ws2.getCell("BW"+[bz]).numFmt = "0.00%"
         }
      }




      if(this.start_h !== ""){
      for(var ax = 0; ax<this.forth_column.length; ax++){
            for(var bz = 1; bz<this.date_use_data.length+4;  bz++){
  
        ws2.getCell(this.forth_column[ax].col_name+bz).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
      
       ws2.getCell(this.forth_column[ax].col_name+bz).alignment = {
          horizontal: "center",
          vertical: "middle",
        };

        ws2.getCell(this.forth_column[ax].col_name+bz).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      }
      }
      }else{
        for(var ax = 0; ax<this.forth_column.length; ax++){
            for(var bz = 1; bz<this.date_use_data2.length+4;  bz++){
  
        ws2.getCell(this.forth_column[ax].col_name+bz).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
      
       ws2.getCell(this.forth_column[ax].col_name+bz).alignment = {
          horizontal: "center",
          vertical: "middle",
        };

        ws2.getCell(this.forth_column[ax].col_name+bz).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      }
      }
      }
      for(var ax = 0; ax<this.fix_week; ax++){
        ws2.getCell("AF"+[ax+4]).value = 1.50/100
        ws2.getCell("AF"+[ax+4]).numFmt = "0.00%"

      }

      console.log(this.row_result_t13)
      console.log(this.fix_week)
      for(var ax = 0; ax<this.fix_week; ax++){
        if(this.start_h !== ""){
         ws2.getCell("A"+[ax+4]).value = this.date_use_data[ax].fmt_first_date + "-" + this.date_use_data[ax].fmt_sec_date
        }else{
         ws2.getCell("A"+[ax+4]).value = this.date_use_data2[ax].first_date + "-" + this.date_use_data2[ax].sec_date
        }
         if(this.row_result_t1[ax].value == 0){
           ws2.getCell("B"+[ax+4]).value = ""
         }else{
         ws2.getCell("B"+[ax+4]).value = this.row_result_t1[ax].value/100
         ws2.getCell("B"+[ax+4]).numFmt = "0.00%"
         }
        
       
         if(this.row_result_t2[ax].value == 0){
           ws2.getCell("C"+[ax+4]).value = ""
         }else{
           ws2.getCell("C"+[ax+4]).value = this.row_result_t2[ax].value/100
         ws2.getCell("C"+[ax+4]).numFmt = "0.00%"
         }
       
       
         if(this.row_result_t3[ax].value == 0){
           ws2.getCell("D"+[ax+4]).value = ""
         }else{
          ws2.getCell("D"+[ax+4]).value = this.row_result_t3[ax].value/100
         ws2.getCell("D"+[ax+4]).numFmt = "0.00%" 
         }

         if(this.row_result_t4[ax].value == 0){
           ws2.getCell("E"+[ax+4]).value = ""
         }else{
         ws2.getCell("E"+[ax+4]).value = this.row_result_t4[ax].value/100
         ws2.getCell("E"+[ax+4]).numFmt = "0.00%"
         }
         
         
         if(this.row_result_t5[ax].value == 0){
        ws2.getCell("F"+[ax+4]).value = ""
         }else{
        ws2.getCell("F"+[ax+4]).value = this.row_result_t5[ax].value/100
         ws2.getCell("F"+[ax+4]).numFmt = "0.00%"
         }


         if(this.row_result_t6[ax].value == 0){
        ws2.getCell("G"+[ax+4]).value = ""
         }else{
        ws2.getCell("G"+[ax+4]).value = this.row_result_t6[ax].value/100
         ws2.getCell("G"+[ax+4]).numFmt = "0.00%"
         }

         
         
         
         if(this.row_result_t7[ax].value == 0){
          ws2.getCell("H"+[ax+4]).value = ""
         }else{
        ws2.getCell("H"+[ax+4]).value = this.row_result_t7[ax].value/100
         ws2.getCell("H"+[ax+4]).numFmt = "0.00%" 
         }

         


         if(this.row_result_t8[ax].value == 0){
           ws2.getCell("I"+[ax+4]).value = ""
         }else{
         ws2.getCell("I"+[ax+4]).value = this.row_result_t8[ax].value/100    
         ws2.getCell("I"+[ax+4]).numFmt = "0.00%"
         }

         
         
         
         if(this.row_result_t9[ax].value == 0){
           ws2.getCell("J"+[ax+4]).value = ""
         }else{
         ws2.getCell("J"+[ax+4]).value = this.row_result_t9[ax].value/100
         ws2.getCell("J"+[ax+4]).numFmt = "0.00%"
         }

         


         if(this.row_result_t10[ax].value == 0){
           ws2.getCell("K"+[ax+4]).value = ""
         }else{
         ws2.getCell("K"+[ax+4]).value = this.row_result_t10[ax].value/100
         ws2.getCell("K"+[ax+4]).numFmt = "0.00%" 
         }

         


         if(this.row_result_t11[ax].value == 0){
           ws2.getCell("L"+[ax+4]).value = ""
         }else{
         ws2.getCell("L"+[ax+4]).value = this.row_result_t11[ax].value/100
         ws2.getCell("L"+[ax+4]).numFmt = "0.00%" 
         }

         


         if(this.row_result_t12[ax].value == 0){
           ws2.getCell("M"+[ax+4]).value = ""
         }else{
          ws2.getCell("M"+[ax+4]).value = this.row_result_t12[ax].value/100
         ws2.getCell("M"+[ax+4]).numFmt = "0.00%" 
         }
        
         
          if(this.row_result_t13[ax].value == 0){
           ws2.getCell("N"+[ax+4]).value = ""
         }else{
          ws2.getCell("N"+[ax+4]).value = this.row_result_t13[ax].value/100
         ws2.getCell("N"+[ax+4]).numFmt = "0.00%" 
         } 

         if(this.row_result_t14[ax].value == 0){
           ws2.getCell("O"+[ax+4]).value = ""
         }else{
          ws2.getCell("O"+[ax+4]).value = this.row_result_t14[ax].value/100
         ws2.getCell("O"+[ax+4]).numFmt = "0.00%" 
         }

         if(this.row_result_t15[ax].value == 0){
           ws2.getCell("P"+[ax+4]).value = ""
         }else{
          ws2.getCell("P"+[ax+4]).value = this.row_result_t15[ax].value/100
         ws2.getCell("P"+[ax+4]).numFmt = "0.00%" 
         }

         if(this.row_result_t16[ax].value == 0){
           ws2.getCell("Q"+[ax+4]).value = ""
         }else{
          ws2.getCell("Q"+[ax+4]).value = this.row_result_t16[ax].value/100
         ws2.getCell("Q"+[ax+4]).numFmt = "0.00%" 
         }

         if(this.row_result_t17[ax].value == 0){
           ws2.getCell("R"+[ax+4]).value = ""
         }else{
          ws2.getCell("R"+[ax+4]).value = this.row_result_t17[ax].value/100
         ws2.getCell("R"+[ax+4]).numFmt = "0.00%" 
         }

         if(this.row_result_t18[ax].value == 0){
           ws2.getCell("S"+[ax+4]).value = ""
         }else{
          ws2.getCell("S"+[ax+4]).value = this.row_result_t18[ax].value/100
         ws2.getCell("S"+[ax+4]).numFmt = "0.00%" 
         }

         if(this.row_result_t19[ax].value == 0){
           ws2.getCell("T"+[ax+4]).value = ""
         }else{
          ws2.getCell("T"+[ax+4]).value = this.row_result_t19[ax].value/100
         ws2.getCell("T"+[ax+4]).numFmt = "0.00%" 
         }

         if(this.row_result_t20[ax].value == 0){
           ws2.getCell("U"+[ax+4]).value = ""
         }else{
          ws2.getCell("U"+[ax+4]).value = this.row_result_t20[ax].value/100
         ws2.getCell("U"+[ax+4]).numFmt = "0.00%" 
         }

         if(this.row_result_tM1[ax].value == 0){
           ws2.getCell("V"+[ax+4]).value = ""
         }else{
         ws2.getCell("V"+[ax+4]).value = this.row_result_tM1[ax].value/100
         ws2.getCell("V"+[ax+4]).numFmt = "0.00%"
         }


         if(this.row_result_tM2[ax].value == 0){
           ws2.getCell("W"+[ax+4]).value = ""
         }else{
          ws2.getCell("W"+[ax+4]).value = this.row_result_tM2[ax].value/100
         ws2.getCell("W"+[ax+4]).numFmt = "0.00%"
         }

         
         
         
         if(this.row_result_tM3[ax].value == 0){
           ws2.getCell("X"+[ax+4]).value = ""
         }else{
           ws2.getCell("X"+[ax+4]).value = this.row_result_tM3[ax].value/100
           ws2.getCell("X"+[ax+4]).numFmt = "0.00%"

         }
 
         
         if(this.row_result_tM4[ax].value == 0){
           ws2.getCell("Y"+[ax+4]).value = ""
         }else{
          ws2.getCell("Y"+[ax+4]).value = this.row_result_tM4[ax].value/100
         ws2.getCell("Y"+[ax+4]).numFmt = "0.00%"
         }

        


         if(this.row_result_tM5[ax].value == 0){
           ws2.getCell("Z"+[ax+4]).value = ""
         }else{
          ws2.getCell("Z"+[ax+4]).value = this.row_result_tM5[ax].value/100
         ws2.getCell("Z"+[ax+4]).numFmt = "0.00%"
         }

         


         if(this.row_result_tM6[ax].value == 0){
          ws2.getCell("AA"+[ax+4]).value = ""
         }else{
         ws2.getCell("AA"+[ax+4]).value = this.row_result_tM6[ax].value/100
         ws2.getCell("AA"+[ax+4]).numFmt = "0.00%"
         }

         if(this.row_result_tM7[ax].value == 0){
          ws2.getCell("AB"+[ax+4]).value = ""
         }else{
         ws2.getCell("AB"+[ax+4]).value = this.row_result_tM7[ax].value/100
         ws2.getCell("AB"+[ax+4]).numFmt = "0.00%"
         }

         if(this.row_result_tM8[ax].value == 0){
          ws2.getCell("AC"+[ax+4]).value = ""
         }else{
         ws2.getCell("AC"+[ax+4]).value = this.row_result_tM8[ax].value/100
         ws2.getCell("AC"+[ax+4]).numFmt = "0.00%"
         }

         if(this.row_result_tM9[ax].value == 0){
          ws2.getCell("AD"+[ax+4]).value = ""
         }else{
         ws2.getCell("AD"+[ax+4]).value = this.row_result_tM9[ax].value/100
         ws2.getCell("AD"+[ax+4]).numFmt = "0.00%"
         }

         if(this.row_result_tM10[ax].value == 0){
          ws2.getCell("AE"+[ax+4]).value = ""
         }else{
         ws2.getCell("AE"+[ax+4]).value = this.row_result_tM10[ax].value/100
         ws2.getCell("AE"+[ax+4]).numFmt = "0.00%"
         }



       
        if(this.row_result_endloss_t1[ax].value == 0){
        ws2.getCell("AG"+[ax+4]).value = ""
        }else{
         ws2.getCell("AG"+[ax+4]).value = this.row_result_endloss_t1[ax].value/100
         ws2.getCell("AG"+[ax+4]).numFmt = "0.00%"
        }
       
        if(this.row_result_endloss_t2[ax].value ==0){
        ws2.getCell("AH"+[ax+4]).value = ""
        }else{
         ws2.getCell("AH"+[ax+4]).value = this.row_result_endloss_t2[ax].value/100
         ws2.getCell("AH"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_t3[ax].value ==0){
        ws2.getCell("AI"+[ax+4]).value = ""
        }else{
         ws2.getCell("AI"+[ax+4]).value = this.row_result_endloss_t3[ax].value/100
         ws2.getCell("AI"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_t4[ax].value == 0){
        ws2.getCell("AJ"+[ax+4]).value = ""
        }else{
         ws2.getCell("AJ"+[ax+4]).value = this.row_result_endloss_t4[ax].value/100
         ws2.getCell("AJ"+[ax+4]).numFmt = "0.00%"
        }

        
        if(this.row_result_endloss_t5[ax].value == 0){
        ws2.getCell("AK"+[ax+4]).value = ""
        }else{
         ws2.getCell("AK"+[ax+4]).value = this.row_result_endloss_t5[ax].value/100
         ws2.getCell("AK"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_t6[ax].value ==0){
        ws2.getCell("AL"+[ax+4]).value = ""
        }else{
         ws2.getCell("AL"+[ax+4]).value = this.row_result_endloss_t6[ax].value/100
         ws2.getCell("AL"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_t7[ax].value == 0){
        ws2.getCell("AM"+[ax+4]).value = ""
        }else{
         ws2.getCell("AM"+[ax+4]).value = this.row_result_endloss_t7[ax].value/100
         ws2.getCell("AM"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_t8[ax].value == 0){
        ws2.getCell("AN"+[ax+4]).value = ""
        }else{
         ws2.getCell("AN"+[ax+4]).value = this.row_result_endloss_t8[ax].value/100
         ws2.getCell("AN"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_t9[ax].value ==0){
        ws2.getCell("AO"+[ax+4]).value = ""
        }else{
         ws2.getCell("AO"+[ax+4]).value = this.row_result_endloss_t9[ax].value/100
         ws2.getCell("AO"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_t10[ax].value == 0){
        ws2.getCell("AP"+[ax+4]).value = ""
        }else{
         ws2.getCell("AP"+[ax+4]).value = this.row_result_endloss_t10[ax].value/100
         ws2.getCell("AP"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_t11[ax].value == 0){
        ws2.getCell("AQ"+[ax+4]).value = ""
        }else{
         ws2.getCell("AQ"+[ax+4]).value = this.row_result_endloss_t11[ax].value/100
         ws2.getCell("AQ"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_t12[ax].value == 0){
        ws2.getCell("AR"+[ax+4]).value = ""
        }else{
         ws2.getCell("AR"+[ax+4]).value = this.row_result_endloss_t12[ax].value/100
         ws2.getCell("AR"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_t13[ax].value == 0){
        ws2.getCell("AS"+[ax+4]).value = ""
        }else{
         ws2.getCell("AS"+[ax+4]).value = this.row_result_endloss_t13[ax].value/100
         ws2.getCell("AS"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_t14[ax].value == 0){
        ws2.getCell("AT"+[ax+4]).value = ""
        }else{
         ws2.getCell("AT"+[ax+4]).value = this.row_result_endloss_t14[ax].value/100
         ws2.getCell("AT"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_t15[ax].value == 0){
        ws2.getCell("AU"+[ax+4]).value = ""
        }else{
         ws2.getCell("AU"+[ax+4]).value = this.row_result_endloss_t15[ax].value/100
         ws2.getCell("AU"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_t16[ax].value == 0){
        ws2.getCell("AV"+[ax+4]).value = ""
        }else{
         ws2.getCell("AV"+[ax+4]).value = this.row_result_endloss_t16[ax].value/100
         ws2.getCell("AV"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_t17[ax].value == 0){
        ws2.getCell("AW"+[ax+4]).value = ""
        }else{
         ws2.getCell("AW"+[ax+4]).value = this.row_result_endloss_t17[ax].value/100
         ws2.getCell("AW"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_t18[ax].value == 0){
        ws2.getCell("AX"+[ax+4]).value = ""
        }else{
         ws2.getCell("AX"+[ax+4]).value = this.row_result_endloss_t18[ax].value/100
         ws2.getCell("AX"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_t19[ax].value == 0){
        ws2.getCell("AY"+[ax+4]).value = ""
        }else{
         ws2.getCell("AY"+[ax+4]).value = this.row_result_endloss_t19[ax].value/100
         ws2.getCell("AY"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_t20[ax].value == 0){
        ws2.getCell("AZ"+[ax+4]).value = ""
        }else{
         ws2.getCell("AZ"+[ax+4]).value = this.row_result_endloss_t20[ax].value/100
         ws2.getCell("AZ"+[ax+4]).numFmt = "0.00%"
        }


        if(this.row_result_endloss_tM1[ax].value == 0){
        ws2.getCell("BA"+[ax+4]).value = ""
        }else{
         ws2.getCell("BA"+[ax+4]).value = this.row_result_endloss_tM1[ax].value/100
         ws2.getCell("BA"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_tM2[ax].value ==0){
        ws2.getCell("BB"+[ax+4]).value = ""
        }else{
         ws2.getCell("BB"+[ax+4]).value = this.row_result_endloss_tM2[ax].value/100
         ws2.getCell("BB"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_tM3[ax].value ==0){
        ws2.getCell("BC"+[ax+4]).value = ""
        }else{
         ws2.getCell("BC"+[ax+4]).value = this.row_result_endloss_tM3[ax].value/100
         ws2.getCell("BC"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_tM4[ax].value ==0){
        ws2.getCell("BD"+[ax+4]).value = ""
        }else{
         ws2.getCell("BD"+[ax+4]).value = this.row_result_endloss_tM4[ax].value/100
         ws2.getCell("BD"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_tM5[ax].value ==0){
        ws2.getCell("BE"+[ax+4]).value = ""
        }else{
         ws2.getCell("BE"+[ax+4]).value = this.row_result_endloss_tM5[ax].value/100
         ws2.getCell("BE"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_tM6[ax].value ==0){
        ws2.getCell("BF"+[ax+4]).value = ""
        }else{
         ws2.getCell("BF"+[ax+4]).value = this.row_result_endloss_tM6[ax].value/100
         ws2.getCell("BF"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_tM7[ax].value ==0){
        ws2.getCell("BG"+[ax+4]).value = ""
        }else{
         ws2.getCell("BG"+[ax+4]).value = this.row_result_endloss_tM7[ax].value/100
         ws2.getCell("BG"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_tM8[ax].value ==0){
        ws2.getCell("BH"+[ax+4]).value = ""
        }else{
         ws2.getCell("BH"+[ax+4]).value = this.row_result_endloss_tM8[ax].value/100
         ws2.getCell("BH"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_tM9[ax].value ==0){
        ws2.getCell("BI"+[ax+4]).value = ""
        }else{
         ws2.getCell("BI"+[ax+4]).value = this.row_result_endloss_tM9[ax].value/100
         ws2.getCell("BI"+[ax+4]).numFmt = "0.00%"
        }

        if(this.row_result_endloss_tM10[ax].value ==0){
        ws2.getCell("BJ"+[ax+4]).value = ""
        }else{
         ws2.getCell("BJ"+[ax+4]).value = this.row_result_endloss_tM10[ax].value/100
         ws2.getCell("BJ"+[ax+4]).numFmt = "0.00%"
        }

          
      }
  
      for(var ax = 0; ax<this.result_actual_end.length; ax++){
        if(this.result_actual_end[ax].value > 0 || isNaN(this.result_actual_end)==false || isFinite(this.result_actual_end)==true){
         ws2.getCell("BL"+[ax+4]).value = this.result_actual_end[ax].value/100
         ws2.getCell("BL"+[ax+4]).numFmt = "0.00%"
        }else{
         ws2.getCell("BL"+[ax+4]).value = 0/100
         ws2.getCell("BL"+[ax+4]).numFmt = "0.00%"
        }
      }

      for(var ax = 0; ax<this.result_actual_splice.length; ax++){
        if(this.result_actual_splice[ax].value > 0 || isNaN(this.result_actual_splice)==false || isFinite(this.result_actual_splice)==true){
         ws2.getCell("BM"+[ax+4]).value = this.result_actual_splice[ax].value/100
         ws2.getCell("BM"+[ax+4]).numFmt = "0.00%"
        }else{
         ws2.getCell("BM"+[ax+4]).value = 0/100
         ws2.getCell("BM"+[ax+4]).numFmt = "0.00%"
        }
      }

      for(var ax = 0; ax<this.result_actual_cut.length; ax++){
        if(this.result_actual_cut[ax].value > 0 || isNaN(this.result_actual_cut)==false || isFinite(this.result_actual_cut)==true){
         ws2.getCell("BN"+[ax+4]).value = this.result_actual_cut[ax].value/100
         ws2.getCell("BN"+[ax+4]).numFmt = "0.00%"
        }else{
         ws2.getCell("BN"+[ax+4]).value = 0/100
         ws2.getCell("BN"+[ax+4]).numFmt = "0.00%"
        }
      }

      for(var ax = 0; ax<this.result_actual_rem.length; ax++){
        if(this.result_actual_rem[ax].value > 0 || isNaN(this.result_actual_rem)==false || isFinite(this.result_actual_rem)==true){
         ws2.getCell("BO"+[ax+4]).value = this.result_actual_rem[ax].value/100
         ws2.getCell("BO"+[ax+4]).numFmt = "0.00%"
        }else{
         ws2.getCell("BO"+[ax+4]).value = 0/100
         ws2.getCell("BO"+[ax+4]).numFmt = "0.00%"
        }
      }

        for(var ax = 0; ax<this.result_actual_end.length; ax++){
      this.sum_actual_all_value = 0
      this.sum_actual_all_value = Number(this.result_actual_end[ax].value) + Number(this.result_actual_splice[ax].value) 
      + Number(this.result_actual_cut[ax].value) + Number(this.result_actual_rem[ax].value)
      if(this.sum_actual_all_value > 0 || isNaN(this.sum_actual_all_value)==false || isFinite(this.sum_actual_all_value)==true){
         ws2.getCell("BP"+[ax+4]).value = this.sum_actual_all_value/100
         ws2.getCell("BP"+[ax+4]).numFmt = "0.00%"

         this.actual_total_length_loss.push({
           value:this.sum_actual_all_value
         })
      }else{
        ws2.getCell("BP"+[ax+4]).value = 0/100
        ws2.getCell("BP"+[ax+4]).numFmt = "0.00%"
        this.actual_total_length_loss.push({
           value:0
         })
      } 
    }

    for(var ax =0; ax<this.actual_width_loss.length; ax++){
     
      if(this.actual_width_loss[ax].value > 0 || isNaN(this.actual_width_loss[ax].value)==false){
        ws2.getCell("BQ"+[ax+4]).value = this.actual_width_loss[ax].value/100
        ws2.getCell("BQ"+[ax+4]).numFmt = "0.00%"
      }else{
        ws2.getCell("BQ"+[ax+4]).value = 0/100
        ws2.getCell("BQ"+[ax+4]).numFmt = "0.00%"
      }
    }
      
     this.$q.loading.hide({
          
        })

   
       workbook.xlsx.writeBuffer().then((data) => {
      const blob = new Blob([data], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8",
        });
        saveAs(blob, "Report Weekly Analysis.xlsx");   
      })   
      //const sheet = workbook.addWorksheet('My Sheet', {views: [{showGridLines: false}]}); ***test show line grid disabled
    

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

        },
       
  },

}

</script>

<style>
.text-danger{
  
  color:red;
  font-size : 1em;
  font-family: sans-serif;

}
</style>