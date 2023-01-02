<template>
  <div class="row">
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

   
    <q-select
        filled
        v-model="multiple_chose"
        :options="options"
        label="Chose Report"
        style="width: 200px"
      />
       <q-select
            
            class="q-px-xl"
            v-model="org"
            :options="option_org"
            style="width:200px"
            disable
            
            label="Chose Org"
          />
    
             <q-btn
      size="md"
      dense
      class="q-px-xl q-py-xs "
      color="primary"
      label="Export"
      @click="exportexcel()"
    />
  </div>
</template>

<script>
import {ref} from 'vue'
import {date} from 'quasar'
import axios from 'axios'
import Detail from '../components/Detail.vue'
import { useQuasar, QSpinnerGears } from 'quasar'
import { onBeforeUnmount } from 'vue'
import * as Excel from "exceljs";
import { saveAs } from "file-saver";
export default {
  data(){
    return{
       multiple_chose: "ALL",
       options:['ALL','Woven','Light-Middle Weight','Fleece','Special'],
        org:'',
        start:'',
        end:'',
        start_date:'',
        multiple:[],
        end_date:'',
          x:0,
        rowexport:[],
      
        end_loss_plug12:[],
        keep_ttlmarker:[],
        keep_avg_end:[],
        keep_splice_Loss:[],
        keep_cut_out_loss:[],
        keep_remnants:[],
        keep_total_loss:[],
        keep_plug12_inch:[],
        keep_line_l:[],
        keep_avg_loss_plug12:[],
        keep_avg_length_yd:[],
        keep_sum_length_splice_loss:[],
        keep_sum_cut_loss_out:[],
        keep_perplug12:[],
        keep_splice_loss_per:[],
        keep_cut_out_loss_per:[],
        keep_remnants_per:[],
        total_overall:[],
        keep_total_percent:[],
        keep_total_rem_percent:[],
        keep_total_loss_per:[],
        keep_total_total_length_percent:[],
        over_all_end_loss:[],
        over_all_splice_loss:[],
        over_all_cut_out_loss:[],
        over_all_remnants:[],
        over_all_total_loss:[],
        check_value:'',
        keep_actual:[],
        export_out:[],
        keep_multi:[],
        rowexport2:[],
        end_loss_type:[],
        splice_loss_type:[],
        cut_out_loss_type:[],
        rem_loss_type:[],
        all_value_before_diff:[],
        all_value_before_target:[],
        result_minus_diff:[],
      
        end_loss_value:[],
        splice_value:[],
        cut_out_value:[],
        remnants_loss_value:[],
        total_length_value:[]


      }
    
  },
  mounted() {
    this.rowx=[]
     let org = this.$q.localStorage.getItem('org')
      this.org = org
         
      axios.get(this.$api_url + "/mainso.php/get_target")
      .then((resp)=>{
      
        this.rowx = resp.data.data

   
        this.end_loss_value.push({
          p1:parseFloat(this.rowx[0].END_LOSS_WOVEN).toFixed(2),
          p2:parseFloat(this.rowx[0].END_LOSS_LIGHT).toFixed(2),
          p3:parseFloat(this.rowx[0].END_LOSS_FLEECE).toFixed(2),
          p4:parseFloat(this.rowx[0].END_LOSS_SPECIAL).toFixed(2)
        })

        this.splice_value.push({
          p1:parseFloat(this.rowx[0].SPLICE_LOSS_WOVEN).toFixed(2),
          p2:parseFloat(this.rowx[0].SPLICE_LOSS_LIGHT).toFixed(2),
          p3:parseFloat(this.rowx[0].SPLICE_LOSS_FLEECE).toFixed(2),
          p4:parseFloat(this.rowx[0].SPLICE_LOSS_SPECIAL).toFixed(2)
        })

        this.cut_out_value.push({
          p1:parseFloat(this.rowx[0].CUT_LOSS_WOVEN).toFixed(2),
          p2:parseFloat(this.rowx[0].CUT_LOSS_LIGHT).toFixed(2),
          p3:parseFloat(this.rowx[0].CUT_LOSS_FLEECE).toFixed(2),
          p4:parseFloat(this.rowx[0].CUT_LOSS_SPECIAL).toFixed(2)
        })

        this.remnants_loss_value.push({
          p1:parseFloat(this.rowx[0].REM_LOSS_WOVEN).toFixed(2),
          p2:parseFloat(this.rowx[0].REM_LOSS_LIGHT).toFixed(2),
          p3:parseFloat(this.rowx[0].REM_LOSS_FLEECE).toFixed(2),
          p4:parseFloat(this.rowx[0].REM_LOSS_SPECIAL).toFixed(2)
        })

        
        this.total_woven = Number(this.rowx[0].END_LOSS_WOVEN) + Number(this.rowx[0].SPLICE_LOSS_WOVEN) + Number(this.rowx[0].CUT_LOSS_WOVEN)
        +Number(this.rowx[0].REM_LOSS_WOVEN)

        this.total_light = Number(this.rowx[0].END_LOSS_LIGHT) + Number(this.rowx[0].SPLICE_LOSS_LIGHT) + Number(this.rowx[0].CUT_LOSS_LIGHT)
        +Number(this.rowx[0].REM_LOSS_LIGHT)

        this.total_fleece = Number(this.rowx[0].END_LOSS_FLEECE) + Number(this.rowx[0].SPLICE_LOSS_FLEECE) + Number(this.rowx[0].CUT_LOSS_FLEECE)
        +Number(this.rowx[0].REM_LOSS_FLEECE)

        this.total_special = Number(this.rowx[0].END_LOSS_SPECIAL) + Number(this.rowx[0].SPLICE_LOSS_SPECIAL) + Number(this.rowx[0].CUT_LOSS_SPECIAL)
        +Number(this.rowx[0].REM_LOSS_SPECIAL)

        this.total_woven = parseFloat(this.total_woven).toFixed(2)
        this.total_light = parseFloat(this.total_light).toFixed(2)
        this.total_fleece = parseFloat(this.total_fleece).toFixed(2)
        this.total_special = parseFloat(this.total_special).toFixed(2)

        this.total_length_value.push({
          p1:this.total_woven,
          p2:this.total_light,
          p3:this.total_fleece,
          p4:this.total_special
        })
        
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
      
  },
    methods:{
      checkdata(){
        
        
          for(var x = 0; x<this.multiple.length; x++){
            
          }
      },
      async  exportexcel(){
        this.rowexport=[],
        this.end_loss_plug12=[],
        this.multiple=[],
        this.keep_ttlmarker=[],
        this.keep_avg_end=[],
        this.keep_splice_Loss=[],
        this.keep_cut_out_loss=[],
        this.keep_remnants=[],
        this.keep_total_loss=[],
        this.keep_plug12_inch=[],
        this.keep_line_l=[],
        this.keep_avg_loss_plug12=[],
        this.keep_avg_length_yd=[],
        this.keep_sum_length_splice_loss=[],
        this.keep_sum_cut_loss_out=[],
        this.keep_perplug12=[],
        this.keep_splice_loss_per=[],
        this.keep_cut_out_loss_per=[],
        this.keep_remnants_per=[],
        this.total_overall=[],
        this.keep_total_percent=[],
        this.keep_total_rem_percent=[],
        this.keep_total_loss_per=[],
        this.keep_total_total_length_percent=[],
        this.over_all_end_loss=[],
        this.over_all_splice_loss=[],
        this.over_all_cut_out_loss=[],
        this.over_all_remnants=[],
        this.over_all_total_loss=[],
        this.check_value='',
        this.keep_actual=[],
        this.export_out=[],
        this.keep_multi=[],
        this.rowexport2=[],
        this.end_loss_type=[],
        this.splice_loss_type=[],
        this.cut_out_loss_type=[],
        this.rem_loss_type=[]
        this.all_value_before_diff=[],
        this.all_value_before_target=[],
        this.result_minus_diff=[]

     
      var x = 0
          const workbook = new Excel.Workbook();
      workbook.creator = "Nanyang";
      workbook.lastModifiedBy = "Admin";
      workbook.created = new Date(2021, 8, 30);
      workbook.modified = new Date();
      workbook.lastPrinted = new Date(2021, 7, 27);
      const worksheet = workbook.addWorksheet("New Sheet");

       worksheet.columns = [ {key: 'A', width: 10},{key: 'B', width: 7.63},{key: 'C', width: 13},{key: 'D', width: 40},{key: 'E', width: 30}
       ,{key: 'F', width: 70},{key: 'G', width: 8.7},{key: 'H', width: 6.7},{key: 'I', width: 6.7},{key: 'J', width: 12},{key: 'K', width: 12},
       {key: 'L', width: 6.7},{key: 'M', width: 6.7},{key: 'N', width: 6.7},{key: 'O', width: 6.7},{key: 'P', width: 6.7},{key: 'Q', width: 6.7},
        {key: 'R', width: 6.7},{key: 'S', width: 12},{key: 'T', width: 6.7},{key: 'U', width: 6.7},{key: 'V', width: 6.7},{key: 'W', width: 6.7},{key: 'X', width: 6.7}];


      let org = this.$q.localStorage.getItem("org");
       worksheet.getCell("A1").value = "SPREAD LOSS AUDIT DAILY REPORT  ( MU-SLA-1 ) : NY"+org //ส่วนหัว
       worksheet.getCell("A1").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet.getCell('A1').font = {
        name: 'Arial Black',
        color: { argb: 'FF0120E5' },
        family: 4,
        size: 20,
        bold: true
      };

      worksheet.getCell("A3").value = "Date"; //ส่วนหัว
       worksheet.getCell("A3").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

       worksheet.getCell("B3").value = this.start_date;
      worksheet.mergeCells("B3:E3");
      worksheet.getCell("B3").alignment = { horizontal: "left",vertical:"middle" };
      worksheet.getCell("B3").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      worksheet.getCell("F3").value = "Fabric Type";
      worksheet.mergeCells("F3:G3");
      worksheet.getCell("F3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("F3").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      worksheet.getCell("H3").value = "Marker";
      worksheet.mergeCells("H3:J3");
      worksheet.getCell("H3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H3").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      worksheet.getCell("K3").value = "Width Loss";
      worksheet.mergeCells("K3:L3");
      worksheet.getCell("K3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K3").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      worksheet.getCell("M3").value = "AVG. End Loss";
      worksheet.mergeCells("M3:O3");
      worksheet.getCell("M3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M3").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      worksheet.getCell("P3").value = "Splice Loss";
      worksheet.mergeCells("P3:Q3");
      worksheet.getCell("P3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P3").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      worksheet.getCell("R3").value = "Cut Out Loss";
      worksheet.mergeCells("R3:S3");
      worksheet.getCell("R3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R3").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      worksheet.getCell("T3").value = "Remnants Loss";
      worksheet.mergeCells("T3:U3");
      worksheet.getCell("T3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T3").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      worksheet.getCell("V3").value = "Total Length Loss";
      worksheet.mergeCells("V3:W3");
      worksheet.getCell("V3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V3").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      if(this.multiple_chose == "ALL"){
        this.multiple_chose = ['Woven','Light-Middle Weight','Fleece','Special']

      }
      const params20 = new FormData();
      for(var ax = 0; ax<this.multiple_chose.length; ax++){
        console.log(this.multiple_chose[ax])
        if(this.multiple_chose[ax] == "Woven"){
          this.multiple_send = 1
        }
        if(this.multiple_chose[ax] == "Light-Middle Weight"){
          this.multiple_send = 2
        }
        if(this.multiple_chose[ax] == "Fleece"){
          this.multiple_send = 3
        }
        if(this.multiple_chose[ax] == "Special"){
          this.multiple_send = 4
        }

      console.log(this.multiple_send)
        params20.append("multiple",this.multiple_send)
        params20.append("start",this.start_date)
        params20.append("org", org);
        for (var pair of params20.entries()) {
                  console.log(pair[0] + ", " + pair[1]);
                }
       await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/check_data_have',
        data:params20,
      }).then((resp)=>{
        console.log(resp.data)
         if(resp.data.data.length > 0){
           this.multiple.push(
            resp.data.data[0].F_TYPE
           )
         }
          
      })
        
      }
     
     console.log(this.multiple)
      const params2 = new FormData();
       params2.append("start",this.start_date)
       params2.append("org", org);

      await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data_all_type',
        data:params2,
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
                this.rowexport2.push({
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
        this.result_woven_ttlmarker = 0
       this.result_light_ttlmarker = 0
       this.result_fleece_ttlmarker = 0
       this.result_special_ttlmarker = 0
       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 1){
         
            this.result_woven_ttlmarker = Number(this.result_woven_ttlmarker) + Number(this.rowexport2[ax].ttl_marker)
           
          
          }

      }

      if(this.result_woven_ttlmarker == 0 ){
        this.total_overall.push({
              result_ttl:0
            })
      }else{
         this.total_overall.push({
              result_ttl:this.result_woven_ttlmarker
            })
      } //end

       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 2){
   
            this.result_light_ttlmarker = Number(this.result_light_ttlmarker) + Number(this.rowexport2[ax].ttl_marker)

          }

      }
  
      if(this.result_light_ttlmarker == 0 ){
        this.total_overall.push({
              result_ttl:0
            }) 
      }else{
        this.total_overall.push({
              result_ttl:this.result_light_ttlmarker
            }) 
        
      }//end

       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 3){
   
            this.result_fleece_ttlmarker = Number(this.result_fleece_ttlmarker) + Number(this.rowexport2[ax].ttl_marker)

          } 
      }
      
      if(this.result_fleece_ttlmarker == 0 ){
        this.total_overall.push({
              result_ttl:0
            }) 
      }else{
        this.total_overall.push({
              result_ttl:this.result_fleece_ttlmarker
            }) 
        
      }//end


       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 4){
   
            this.result_special_ttlmarker = Number(this.result_special_ttlmarker) + Number(this.rowexport2[ax].ttl_marker)

          } 
      }
      
      if(this.result_special_ttlmarker == 0 ){
        this.total_overall.push({
              result_ttl:0
            }) 
      }else{
        this.total_overall.push({
              result_ttl:this.result_special_ttlmarker
            }) 
        
      }//end total_overall

      this.end_loss_type_woven = 0
       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 1){
   
            this.end_loss_type_woven = Number(this.end_loss_type_woven) + Number(this.rowexport2[ax].endless_length_yd)

          } 

       }
       
       this.end_loss_type_woven = Number(this.end_loss_type_woven)/Number(this.result_woven_ttlmarker)*100
       
       if(this.end_loss_type_woven == 0 || isNaN(this.end_loss_type_woven)==true){
         this.end_loss_type.push({
           result_endloss:0
         })
       }else{
         this.end_loss_type.push({
           result_endloss:this.end_loss_type_woven
         })
       }
       
        this.end_loss_type_light = 0
       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 2){
   
            this.end_loss_type_light = Number(this.end_loss_type_light) + Number(this.rowexport2[ax].endless_length_yd)

          } 

       }
       this.end_loss_type_light = this.end_loss_type_light / this.result_light_ttlmarker*100
       if(this.end_loss_type_light == 0 || isNaN(this.end_loss_type_light)==true){
         this.end_loss_type.push({
           result_endloss:0
         })
       }else{
         this.end_loss_type.push({
           result_endloss:this.end_loss_type_light
         })
       }

         this.end_loss_type_fleece = 0
       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 3 ){
   
            this.end_loss_type_fleece = Number(this.end_loss_type_fleece) + Number(this.rowexport2[ax].endless_length_yd)

          } 

       }
        this.end_loss_type_fleece = this.end_loss_type_fleece / this.result_fleece_ttlmarker*100
       if(this.end_loss_type_fleece == 0 || isNaN(this.end_loss_type_fleece)==true){
         this.end_loss_type.push({
           result_endloss:0
         })
       }else{
         this.end_loss_type.push({
           result_endloss:this.end_loss_type_fleece
         })
       }


       this.end_loss_type_special = 0
       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 4 ){
   
            this.end_loss_type_special = Number(this.end_loss_type_special) + Number(this.rowexport2[ax].endless_length_yd)

          } 

       }
        this.end_loss_type_special = this.end_loss_type_special / this.result_special_ttlmarker*100
       if(this.end_loss_type_special == 0 || isNaN(this.end_loss_type_special)==true){
         this.end_loss_type.push({
           result_endloss:0
         })
       }else{
         this.end_loss_type.push({
           result_endloss:this.end_loss_type_special
         })
       }

      //end_loss_type 

       this.splice_loss_type_woven = 0
       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 1){
   
            this.splice_loss_type_woven = Number(this.splice_loss_type_woven) + Number(this.rowexport2[ax].spliceloss)

          } 

       }
       
       this.splice_loss_type_woven = Number(this.splice_loss_type_woven)/Number(this.result_woven_ttlmarker)*100
       
       if(this.splice_loss_type_woven == 0 || isNaN(this.splice_loss_type_woven)==true){
         this.splice_loss_type.push({
           result_splice_loss:0
         })
       }else{
         this.splice_loss_type.push({
           result_splice_loss:this.splice_loss_type_woven
         })
       }
       
        this.splice_loss_type_light = 0
       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 2){
   
            this.splice_loss_type_light = Number(this.splice_loss_type_light) + Number(this.rowexport2[ax].spliceloss)

          } 

       }
       this.splice_loss_type_light = this.splice_loss_type_light / this.result_light_ttlmarker*100
       if(this.splice_loss_type_light == 0 || isNaN(this.splice_loss_type_light)==true){
         this.splice_loss_type.push({
           result_splice_loss:0
         })
       }else{
         this.splice_loss_type.push({
           result_splice_loss:this.splice_loss_type_light
         })
       }

         this.splice_loss_type_fleece = 0
       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 3 ){
   
            this.splice_loss_type_fleece = Number(this.splice_loss_type_fleece) + Number(this.rowexport2[ax].spliceloss)

          } 

       }
        this.splice_loss_type_fleece = this.splice_loss_type_fleece / this.result_fleece_ttlmarker*100
       if(this.splice_loss_type_fleece == 0 || isNaN(this.splice_loss_type_fleece)==true){
         this.splice_loss_type.push({
           result_splice_loss:0
         })
       }else{
         this.splice_loss_type.push({
           result_splice_loss:this.splice_loss_type_fleece
         })
       }


       this.splice_loss_type_special = 0
       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 4 ){
   
            this.splice_loss_type_special = Number(this.splice_loss_type_special) + Number(this.rowexport2[ax].spliceloss)

          } 

       }
        this.splice_loss_type_special = this.splice_loss_type_special / this.result_special_ttlmarker*100
       if(this.splice_loss_type_special == 0 || isNaN(this.splice_loss_type_special)==true){
         this.splice_loss_type.push({
           result_splice_loss:0
         })
       }else{
         this.splice_loss_type.push({
           result_splice_loss:this.splice_loss_type_special
         })
       }
      //end_splice_loss

       this.cut_out_loss_type_woven = 0
       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 1){
   
            this.cut_out_loss_type_woven = Number(this.cut_out_loss_type_woven) + Number(this.rowexport2[ax].totalcutout)

          } 

       }
       
       this.cut_out_loss_type_woven = Number(this.cut_out_loss_type_woven)/Number(this.result_woven_ttlmarker)*100
       
       if(this.cut_out_loss_type_woven == 0 || isNaN(this.cut_out_loss_type_woven)==true){
         this.cut_out_loss_type.push({
           result_cut_out_loss:0
         })
       }else{
         this.cut_out_loss_type.push({
           result_cut_out_loss:this.cut_out_loss_type_woven
         })
       }
       
        this.cut_out_loss_type_light = 0
       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 2){
   
            this.cut_out_loss_type_light = Number(this.cut_out_loss_type_light) + Number(this.rowexport2[ax].totalcutout)

          } 

       }
       this.cut_out_loss_type_light = this.cut_out_loss_type_light / this.result_light_ttlmarker*100
       if(this.cut_out_loss_type_light == 0 || isNaN(this.cut_out_loss_type_light)==true){
         this.cut_out_loss_type.push({
           result_cut_out_loss:0
         })
       }else{
         this.cut_out_loss_type.push({
           result_cut_out_loss:this.cut_out_loss_type_light
         })
       }

         this.cut_out_loss_type_fleece = 0
       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 3 ){
   
            this.cut_out_loss_type_fleece = Number(this.cut_out_loss_type_fleece) + Number(this.rowexport2[ax].totalcutout)

          } 

       }
        this.cut_out_loss_type_fleece = this.cut_out_loss_type_fleece / this.result_fleece_ttlmarker*100
       if(this.cut_out_loss_type_fleece == 0 || isNaN(this.cut_out_loss_type_fleece)==true){
         this.cut_out_loss_type.push({
           result_cut_out_loss:0
         })
       }else{
         this.cut_out_loss_type.push({
           result_cut_out_loss:this.cut_out_loss_type_fleece
         })
       }


       this.cut_out_loss_type_special = 0
       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 4 ){
   
            this.cut_out_loss_type_special = Number(this.cut_out_loss_type_special) + Number(this.rowexport2[ax].totalcutout)

          } 

       }
        this.cut_out_loss_type_special = this.cut_out_loss_type_special / this.result_special_ttlmarker*100
       if(this.cut_out_loss_type_special == 0 || isNaN(this.cut_out_loss_type_special)==true){
         this.cut_out_loss_type.push({
           result_cut_out_loss:0
         })
       }else{
         this.cut_out_loss_type.push({
           result_cut_out_loss:this.cut_out_loss_type_special
         })
       }
      //end_cut_out_loss_type


        this.rem_loss_type_woven = 0
       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 1){
   
            this.rem_loss_type_woven = Number(this.rem_loss_type_woven) + Number(this.rowexport2[ax].rement)

          } 

       }
       
       this.rem_loss_type_woven = Number(this.rem_loss_type_woven)/Number(this.result_woven_ttlmarker)*100
       
       if(this.rem_loss_type_woven == 0 || isNaN(this.rem_loss_type_woven)==true){
         this.rem_loss_type.push({
           result_rem_loss:0
         })
       }else{
         this.rem_loss_type.push({
           result_rem_loss:this.rem_loss_type_woven
         })
       }
       
        this.rem_loss_type_light = 0
       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 2){
   
            this.rem_loss_type_light = Number(this.rem_loss_type_light) + Number(this.rowexport2[ax].rement)

          } 

       }
       this.rem_loss_type_light = this.rem_loss_type_light / this.result_light_ttlmarker*100
       if(this.rem_loss_type_light == 0 || isNaN(this.rem_loss_type_light)==true){
         this.rem_loss_type.push({
           result_rem_loss:0
         })
       }else{
         this.rem_loss_type.push({
           result_rem_loss:this.rem_loss_type_light
         })
       }

         this.rem_loss_type_fleece = 0
       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 3 ){
   
            this.rem_loss_type_fleece = Number(this.rem_loss_type_fleece) + Number(this.rowexport2[ax].rement)

          } 

       }
        this.rem_loss_type_fleece = this.rem_loss_type_fleece / this.result_fleece_ttlmarker*100
       if(this.rem_loss_type_fleece == 0 || isNaN(this.rem_loss_type_fleece)==true){
         this.rem_loss_type.push({
           result_rem_loss:0
         })
       }else{
         this.rem_loss_type.push({
           result_rem_loss:this.rem_loss_type_fleece
         })
       }


       this.rem_loss_type_special = 0
       for(var ax = 0; ax<this.rowexport2.length; ax++){
       
          if(this.rowexport2[ax].f_type == 4 ){
   
            this.rem_loss_type_special = Number(this.rem_loss_type_special) + Number(this.rowexport2[ax].rement)

          } 

       }
        this.rem_loss_type_special = this.rem_loss_type_special / this.result_special_ttlmarker*100
       if(this.rem_loss_type_special == 0 || isNaN(this.rem_loss_type_special)==true){
         this.rem_loss_type.push({
           result_rem_loss:0
         })
       }else{
         this.rem_loss_type.push({
           result_rem_loss:this.rem_loss_type_special
         })
       }

        
      

      });//end all

        
      
          var a = 4 // first row head column Date per table
          var b = 5 //mergeA
           var count = 0
       for( x=0; x<this.multiple.length; x++){
         
      
         
          this.keep_multi.push({
         multiple:this.multiple[x]
         })
      
          const params = new FormData()

        params.append("start",this.start_date)
        
        let org = this.$q.localStorage.getItem("org");
        params.append("org", org);

        if(this.multiple[x] == "1"){
          this.chosen = 1
        }
        if(this.multiple[x] == "2"){
          this.chosen = 2
        }
        if(this.multiple[x] == "3"){
          this.chosen = 3
        }
        if(this.multiple[x] == "4"){
          this.chosen = 4
        }
        params.append("f_type",this.chosen)
        for (var pair of params.entries()) {
                  console.log(pair[0] + ", " + pair[1]);
                }

     await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data_2',
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
      
         if(count == 0){
         
       worksheet.getCell("A"+a).value = "Date";
      worksheet.mergeCells("A"+a +":" + "A"+b);
      worksheet.getCell("A"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "A" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].mu_date;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       }
      var ty = this.length
     

      if(this.multiple[x] == "1"){
        this.total_type = "Woven"
      }else if(this.multiple[x] == "2"){
        this.total_type = "Light-Middle Weight"
      }else if(this.multiple[x] == "3"){
        this.total_type = "Fleece"
      }else if(this.multiple[x] == "4"){
        this.total_type = "Special"
      }
      worksheet.getCell("A"+ty).value = "Total Type :"+this.total_type;
      worksheet.mergeCells("A"+ty +":" + "I"+ty);
      worksheet.getCell("A"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
        this.result_ttlmarker = 0
      for(var re_ttlmarker = 0; re_ttlmarker<resp.data.data.length; re_ttlmarker++ ){
        this.result_ttlmarker =  Number(this.result_ttlmarker) + Number(this.rowexport[re_ttlmarker].ttl_marker)
        
        
      }
      
      this.keep_ttlmarker.push({
          result_ttlmarker:this.result_ttlmarker,
        })
     

      

      
        
      this.result_ttlmarker_2 = parseFloat(this.result_ttlmarker).toFixed(2)
        worksheet.getCell("J"+ty).value = Number(this.result_ttlmarker_2);
        worksheet.getCell("J"+ty).numFmt = "#,##0.00"
      worksheet.getCell("J"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      
   this.result_plug12 = 0
       this.result_plug12_sum = 0
      for(var re_plug1_2 = 0; re_plug1_2<resp.data.data.length; re_plug1_2++ ){
        this.result_plug12 = this.rowexport[re_plug1_2].plug12 * this.rowexport[re_plug1_2].ttl_marker
        this.result_plug12_sum = Number(this.result_plug12_sum)+Number(this.result_plug12)
      }
       this.result_plug12_divide = this.result_plug12_sum/this.result_ttlmarker

    
     
     this.result_plug12 = this.result_plug12/resp.data.data.length 
     this.keep_plug12_inch.push({
          result_plug12_inch:this.result_plug12_divide,
        })
        
        
      this.plug12_divide_last = parseFloat(this.result_plug12_divide).toFixed(2)
        worksheet.getCell("K"+ty).value = Number(this.plug12_divide_last);
      worksheet.getCell("K"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 

      
      this.result_perplug12 = 0
      this.result_perplug12_sum = 0

      for(var re_perplug12 = 0; re_perplug12<resp.data.data.length; re_perplug12++ ){
        this.result_perplug12 = this.rowexport[re_perplug12].widthloss * this.rowexport[re_perplug12].ttl_marker
        this.result_perplug12_sum =  Number(this.result_perplug12_sum) + Number(this.result_perplug12)
        
      }
      this.result_perplug12_divide =  this.result_perplug12_sum/this.result_ttlmarker
      this.keep_perplug12.push({
          result_perplug12:this.result_perplug12_divide,
        })

     
     //this.result_perplug12 = this.result_perplug12/resp.data.data.length 

      worksheet.getCell("L"+ty).value = this.result_perplug12_divide/100
      worksheet.getCell("L"+ty).numFmt = '0.00%'; 
      worksheet.getCell("L"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      
      this.end_loss_plug12 = []
       for(var replug12_inch = 0; replug12_inch<resp.data.data.length; replug12_inch++){
      this.result_plug12_inch = (this.rowexport[replug12_inch].endless_length_yd*36)/(this.rowexport[replug12_inch].ttl_marker/this.rowexport[replug12_inch].length_ydb)
    
      this.end_loss_plug12.push(this.result_plug12_inch)
        
      }
        this.result_plug12_avg = 0;
        this.result_plug12_avg_sum = 0;
        
        for(var replug12_avg = 0; replug12_avg<this.end_loss_plug12.length; replug12_avg++){
        this.result_plug12_avg= this.end_loss_plug12[replug12_avg] * this.rowexport[replug12_avg].ttl_marker
        this.result_plug12_avg_sum = Number(this.result_plug12_avg_sum)+Number(this.result_plug12_avg)
        
      }  
     
        this.all_result_avg = this.result_plug12_avg_sum / this.result_ttlmarker
        this.keep_avg_loss_plug12.push({
          all_result_avg:this.all_result_avg
        })   
        
      this.all_result_avg_last = parseFloat(this.all_result_avg).toFixed(2)
      worksheet.getCell("M"+ty).value = Number(this.all_result_avg_last);
      worksheet.getCell("M"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      this.result_length_yd = 0
      for(var re_length_yd = 0; re_length_yd<resp.data.data.length; re_length_yd++ ){
        this.result_length_yd =  Number(this.result_length_yd) + Number(this.rowexport[re_length_yd].endless_length_yd)
        
      }
 
     //this.result_length_yd = this.result_length_yd/resp.data.data.length

     this.keep_avg_length_yd.push({
       all_result_avg_length:this.result_length_yd
     })

     

      this.result_length_yd_last = parseFloat(this.result_length_yd).toFixed(2)
        worksheet.getCell("N"+ty).value = Number(this.result_length_yd_last);
      worksheet.getCell("N"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };


       this.result_avg_per = 0;
      this.result_avg_per = (this.result_length_yd / this.result_ttlmarker)*100
     
      this.keep_avg_end.push({
       result_avg_per: this.result_avg_per
      })
     
      worksheet.getCell("O"+ty).value = this.result_avg_per/100
      worksheet.getCell("O"+ty).numFmt = '0.00%'; 
      worksheet.getCell("O"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      this.result_splice = 0
      for(var re_splice = 0; re_splice<resp.data.data.length; re_splice++ ){
        this.result_splice =  Number(this.result_splice) + Number(this.rowexport[re_splice].spliceloss)
        
      }
     this.sum_result_splice = this.result_splice
     //this.result_splice = this.result_splice/resp.data.data.length
     this.keep_splice_Loss.push({
          result_splice:this.result_splice
     })

      this.result_splice = parseFloat(this.result_splice).toFixed(2)
        worksheet.getCell("P"+ty).value = Number(this.result_splice);
      worksheet.getCell("P"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      
      this.result_splicelossper = 0;
      this.result_splicelossper = this.result_splice / this.result_ttlmarker*100
       this.keep_splice_loss_per.push({
          result_splice_per:this.result_splicelossper
     })
      worksheet.getCell("Q"+ty).value = this.result_splicelossper/100
      worksheet.getCell("Q"+ty).numFmt = '0.00%'; 
      worksheet.getCell("Q"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };


       this.result_cut_out_loss = 0
      for(var re_cut_out = 0; re_cut_out<resp.data.data.length; re_cut_out++ ){
        this.result_cut_out_loss =  Number(this.result_cut_out_loss) + Number(this.rowexport[re_cut_out].totalcutout)
        
      }
     this.sum_cut_loss_out = this.result_cut_out_loss
     //this.result_cut_out_loss = this.result_cut_out_loss/resp.data.data.length 
     this.keep_cut_out_loss.push({
       result_cut_out_loss: this.result_cut_out_loss
     })

      this.result_cut_out_loss_last = parseFloat(this.result_cut_out_loss).toFixed(2)
        worksheet.getCell("R"+ty).value = Number(this.result_cut_out_loss_last);
      worksheet.getCell("R"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      
      this.result_cut_out_loss_per = 0;
      this.result_cut_out_loss_per = this.result_cut_out_loss / this.result_ttlmarker
      this.keep_cut_out_loss_per.push({
        result_cut_out_loss_per:this.result_cut_out_loss_per
      })
      worksheet.getCell("S"+ty).value = this.result_cut_out_loss_per/100
      worksheet.getCell("S"+ty).numFmt = '0.00%'; 
      worksheet.getCell("S"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      this.result_remnants = 0
      for(var re_remnants = 0; re_remnants<resp.data.data.length; re_remnants++ ){
        this.result_remnants =  Number(this.result_remnants) + Number(this.rowexport[re_remnants].rement)
        
      }
     this.sum_remnants = this.result_remnants
    // this.result_remnants = this.result_remnants/resp.data.data.length
     this.keep_remnants.push({
       result_remnants:this.result_remnants
     })

      this.result_remnants_last = parseFloat(this.result_remnants).toFixed(2)
        worksheet.getCell("T"+ty).value = Number(this.result_remnants_last);
      worksheet.getCell("T"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      this.result_remnants_per = 0;
      this.result_remnants_per = (this.result_remnants / this.result_ttlmarker)*100
      this.keep_remnants_per.push({
        remnants_per:this.result_remnants_per
      })
      worksheet.getCell("U"+ty).value = this.result_remnants_per/100
      worksheet.getCell("U"+ty).numFmt = '0.00%'; 
      worksheet.getCell("U"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };


      this.result_total_length = 0
      for(var re_total_length = 0; re_total_length<resp.data.data.length; re_total_length++ ){
      this.result_total_length =  Number(this.result_total_length) + Number(this.rowexport[re_total_length].totallenthloss)
        
      }
     this.sum_total_length = this.result_total_length
     //this.result_total_length = this.result_total_length/resp.data.data.length
     this.keep_total_loss.push({
       result_total_length:this.result_total_length
     }) 

      this.result_total_length_last = parseFloat(this.result_total_length).toFixed(2)
        worksheet.getCell("V"+ty).value = Number(this.result_total_length_last);
      worksheet.getCell("V"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

       this.result_total_length_per = 0;
      this.result_total_length_per = this.result_total_length / this.result_ttlmarker*100
      this.keep_total_loss_per.push({
        result_total_loss_per:this.result_total_length_per
      })
     
      worksheet.getCell("W"+ty).value = this.result_total_length_per/100
      worksheet.getCell("W"+ty).numFmt = '0.00%'; 
      worksheet.getCell("W"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };








       //B---------------------------------------------------------------------
       worksheet.getCell("B"+a).value = "Table";
      worksheet.mergeCells("B"+a +":" + "B"+b);
      worksheet.getCell("B"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("B"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "B" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_table;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       }

        //C------------------------------------------------------------------------
        worksheet.getCell("C"+a).value = "S/O no.";
      worksheet.mergeCells("C"+a +":" + "C"+b);
      worksheet.getCell("C"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("C"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "C" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_no_short;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

        //D------------------------------------------------------------------------
        worksheet.getCell("D"+a).value = "Customer";
      worksheet.mergeCells("D"+a +":" + "D"+b);
      worksheet.getCell("D"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("D"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "D" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].customer;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //E-----------------------------------------------------------
      worksheet.getCell("E"+a).value = "GM08 no.";
      worksheet.mergeCells("E"+a +":" + "E"+b);
      worksheet.getCell("E"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("E"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "E" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_no;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //F-------------------------------------------------------------
       if(this.multiple[x] == 1){
         this.multiple_rec = "Woven"
       }
       if(this.multiple[x] == 2){
         this.multiple_rec = "Light & Middle Weigth"
       }
       if(this.multiple[x] == 3){
         this.multiple_rec = "Fleece"
       }
       if(this.multiple[x] == 4){
         this.multiple_rec = "Special"
       }

       
      worksheet.getCell("F"+a).value = this.multiple_rec;
     worksheet.mergeCells("F"+a +":" + "F"+b);
      worksheet.getCell("F"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("F"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "F" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].fabric_type;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
     
      //G----------------------------------------------------------------------

       worksheet.getCell("G"+a).value = "Observe";
      worksheet.getCell("G"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("G"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("G"+b).value = "Width";
      worksheet.getCell("G"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("G"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "G" + ab
      
            var z = ab-a-2;
       this.last_value_obswidth = parseFloat(this.rowexport[z].obs_width).toFixed(2);
       worksheet.getCell(y).value = Number(this.last_value_obswidth)
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //H-------------------------------------------------------
       worksheet.getCell("H"+a).value = "Width";
      worksheet.getCell("H"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("H"+b).value = "(inch)";
      worksheet.getCell("H"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "H" + ab
      
            var z = ab-a-2;
      this.last_width_inch = parseFloat(this.rowexport[z].width_inch).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_width_inch);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
      //I-----------------------------------------------------------
         worksheet.getCell("I"+a).value = "Length";
      worksheet.getCell("I"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("I"+b).value = "(yd)";
      worksheet.getCell("I"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 

      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "I" + ab
      
            var z = ab-a-2;

      this.last_length_ydb = parseFloat(this.rowexport[z].length_ydb).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_length_ydb);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }


       //J-----------------------------------------------------------
         worksheet.getCell("J"+a).value = "Ttl Marker";
      worksheet.getCell("J"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("J"+b).value = "Length(yd)";
      worksheet.getCell("J"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "J" + ab
      
            var z = ab-a-2;
      this.last_ttl_marker = parseFloat(this.rowexport[z].ttl_marker).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_ttl_marker);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       }
         

//K-------------------------------------------------------------------------------
         worksheet.getCell("K"+a).value = "Plug 1+2";
      worksheet.getCell("K"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("K"+b).value = "(inch)";
      worksheet.getCell("K"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
         this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "K" + ab
      
            var z = ab-a-2;
      this.last_plug12 = parseFloat(this.rowexport[z].plug12).toFixed(2)
      worksheet.getCell(y).value = Number(this.last_plug12);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

//L------------------------------------------------------------------------------------
         worksheet.getCell("L"+a).value = "%";
      worksheet.getCell("L"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("L"+b).value = 1.50/100;
       worksheet.getCell("L"+b).numFmt = '0.00%'; 
      worksheet.getCell("L"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

        this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "L" + ab
      
            var z = ab-a-2;
        this.last_widthloss = parseFloat(this.rowexport[z].widthloss).toFixed(2)
      worksheet.getCell(y).value = this.last_widthloss/100;
      worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
       
  //M--------------------------------------------------------------
        worksheet.getCell("M"+a).value = "Plug 1+2";
      worksheet.getCell("M"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("M"+b).value = "(inch)";
      worksheet.getCell("M"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "M" + ab
      
            var z = ab-a-2;

      this.result_plug12_inch = parseFloat((this.rowexport[z].endless_length_yd*36)/(this.rowexport[z].ttl_marker/this.rowexport[z].length_ydb)).toFixed(2)
      this.last_result_plug12_inch = parseFloat(this.result_plug12_inch).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_result_plug12_inch);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

    //N----------------------------------------------------------
       worksheet.getCell("N"+a).value = "Length";
      worksheet.getCell("N"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("N"+b).value = "(yd)";
      worksheet.getCell("N"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
        this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "N" + ab
      
            var z = ab-a-2;
      this.last_endless_yd = parseFloat(this.rowexport[z].endless_length_yd).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_endless_yd);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //O------------------------------------------------------------------------
        worksheet.getCell("O"+a).value = "%";
      worksheet.getCell("O"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      
      //'Woven','Light-Middle Weight','Fleece','Special']

      if(this.multiple[x] == "Woven"){
        this.check_value = this.end_loss_value[0].p1
      }else if(this.multiple[x] == "Light-Middle Weight"){
        this.check_value = this.end_loss_value[0].p2
      }else if(this.multiple[x] == "Fleece"){
        this.check_value = this.end_loss_value[0].p3
      }else if(this.multiple[x] == "Special"){
        this.check_value = this.end_loss_value[0].p4
      }
      
       worksheet.getCell("O"+b).value = this.check_value/100;
       worksheet.getCell("O"+b).numFmt = '0.00%'; 
      worksheet.getCell("O"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "O" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].avg_end/100;
       worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
      //P--------------------------------------------------------------------------
         worksheet.getCell("P"+a).value = "Length";
      worksheet.getCell("P"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("P"+b).value = "(yd)";
      worksheet.getCell("P"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "P" + ab
      
            var z = ab-a-2;

      this.last_splice_loss = parseFloat(this.rowexport[z].spliceloss).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_splice_loss);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //Q------------------------------------------------------------------------------
         worksheet.getCell("Q"+a).value = "%";
      worksheet.getCell("Q"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 

       if(this.multiple[x] == "Woven"){
        this.check_value = this.splice_value[0].p1
      }else if(this.multiple[x] == "Light-Middle Weight"){
        this.check_value = this.splice_value[0].p2
      }else if(this.multiple[x] == "Fleece"){
        this.check_value = this.splice_value[0].p3
      }else if(this.multiple[x] == "Special"){
        this.check_value = this.splice_value[0].p4
      }

       worksheet.getCell("Q"+b).value = this.check_value/100;
       worksheet.getCell("Q"+b).numFmt = '0.00%'; 
      worksheet.getCell("Q"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "Q" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value =this.rowexport[z].splicelossper/100;
       worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //R------------------------------------------------------------------------------
         worksheet.getCell("R"+a).value = "Length";
      worksheet.getCell("R"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("R"+b).value = "(yd)";
      worksheet.getCell("R"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
          this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "R" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].totalcutout);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }


      //S------------------------------------------------------------------------------
         worksheet.getCell("S"+a).value = "%";
      worksheet.getCell("S"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
        if(this.multiple[x] == "Woven"){
        this.check_value = this.cut_out_value[0].p1
      }else if(this.multiple[x] == "Light-Middle Weight"){
        this.check_value = this.cut_out_value[0].p2
      }else if(this.multiple[x] == "Fleece"){
        this.check_value = this.cut_out_value[0].p3
      }else if(this.multiple[x] == "Special"){
        this.check_value = this.cut_out_value[0].p4
      }
       worksheet.getCell("S"+b).value = this.check_value/100
       worksheet.getCell("S"+b).numFmt = '0.00%'; 
      worksheet.getCell("S"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      
        this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "S" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].totalcutoutper/100
       worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //T-------------------------------------------------------------------------------

         worksheet.getCell("T"+a).value = "Length";
      worksheet.getCell("T"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("T"+b).value = "(yd)";
      worksheet.getCell("T"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "T" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].rement);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //U-------------------------------------------------------------------------------

         worksheet.getCell("U"+a).value = "%";
      worksheet.getCell("U"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
        if(this.multiple[x] == "Woven"){
        this.check_value = this.remnants_loss_value[0].p1
      }else if(this.multiple[x] == "Light-Middle Weight"){
        this.check_value = this.remnants_loss_value[0].p2
      }else if(this.multiple[x] == "Fleece"){
        this.check_value = this.remnants_loss_value[0].p3
      }else if(this.multiple[x] == "Special"){
        this.check_value = this.remnants_loss_value[0].p4
      }
       worksheet.getCell("U"+b).value = this.check_value/100;
       worksheet.getCell("U"+b).numFmt = '0.00%'; 
      worksheet.getCell("U"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "U" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].rement_per/100;
       worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //V-------------------------------------------------------------------------------

         worksheet.getCell("V"+a).value = "Length";
      worksheet.getCell("V"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("V"+b).value = "(yd)";
      worksheet.getCell("V"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "V" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].totallenthloss);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }


       //W-------------------------------------------------------------------------------

         worksheet.getCell("W"+a).value = "%of";
      worksheet.getCell("W"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
        if(this.multiple[x] == "Woven"){
        this.check_value = this.total_length_value[0].p1
      }else if(this.multiple[x] == "Light-Middle Weight"){
        this.check_value = this.total_length_value[0].p2
      }else if(this.multiple[x] == "Fleece"){
        this.check_value = this.total_length_value[0].p3
      }else if(this.multiple[x] == "Special"){
        this.check_value = this.total_length_value[0].p4
      }
       worksheet.getCell("W"+b).value = this.check_value/100;
       worksheet.getCell("W"+b).numFmt = '0.00%'; 
      worksheet.getCell("W"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "W" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].totallenthlossper/100;
       worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
      
        
         } else if(count == 1){
             

             a = this.length+2
             b = this.length+3
     
      worksheet.getCell("A"+a).value = "Date";
      worksheet.mergeCells("A"+a +":" + "A"+b);
      worksheet.getCell("A"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "A" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].mu_date;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
       var ty = this.length+x
      if(this.multiple[x] == "1"){
        this.total_type = "Woven"
      }else if(this.multiple[x] == "2"){
        this.total_type = "Light-Middle Weight"
      }else if(this.multiple[x] == "3"){
        this.total_type = "Fleece"
      }else if(this.multiple[x] == "4"){
        this.total_type = "Special"
      }
       worksheet.getCell("A"+ty).value = "Total Type :"+this.total_type;
      worksheet.mergeCells("A"+ty +":" + "I"+ty);
      worksheet.getCell("A"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+ty).border = {
         top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
        this.result_ttlmarker = 0
      for(var re_ttlmarker = 0; re_ttlmarker<resp.data.data.length; re_ttlmarker++ ){
        this.result_ttlmarker =  Number(this.result_ttlmarker) + Number(this.rowexport[re_ttlmarker].ttl_marker)
      
      }
        this.keep_ttlmarker.push({
          result_ttlmarker:this.result_ttlmarker,
        })
      this.result_ttlmarker_2 = parseFloat(this.result_ttlmarker).toFixed(2)
        worksheet.getCell("J"+ty).value = Number(this.result_ttlmarker_2);
        worksheet.getCell("J"+ty).numFmt = "#,##0.00"
      worksheet.getCell("J"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+ty).border = {
          top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };  
      
     this.result_plug12 = 0
       this.result_plug12_sum = 0
      for(var re_plug1_2 = 0; re_plug1_2<resp.data.data.length; re_plug1_2++ ){
        this.result_plug12 = this.rowexport[re_plug1_2].plug12 * this.rowexport[re_plug1_2].ttl_marker
        this.result_plug12_sum = Number(this.result_plug12_sum)+Number(this.result_plug12)
        
      }
      
       this.result_plug12_divide = this.result_plug12_sum/this.result_ttlmarker


    
     
     this.result_plug12 = this.result_plug12/resp.data.data.length 
     this.keep_plug12_inch.push({
          result_plug12_inch:this.result_plug12_divide,
        })
        
        

        this.plug12_divide_last = parseFloat(this.result_plug12_divide).toFixed(2)
        worksheet.getCell("K"+ty).value = Number(this.plug12_divide_last);
      worksheet.getCell("K"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+ty).border = {
          top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
 this.result_perplug12 = 0
      this.result_perplug12_sum = 0

      for(var re_perplug12 = 0; re_perplug12<resp.data.data.length; re_perplug12++ ){
        this.result_perplug12 = this.rowexport[re_perplug12].widthloss * this.rowexport[re_perplug12].ttl_marker
        this.result_perplug12_sum =  Number(this.result_perplug12_sum) + Number(this.result_perplug12)
        
      }
      this.result_perplug12_divide =  this.result_perplug12_sum/this.result_ttlmarker
      this.keep_perplug12.push({
          result_perplug12:this.result_perplug12_divide,
        })


     //this.result_perplug12_l = this.result_perplug12/resp.data.data.length 
     

        worksheet.getCell("L"+ty).value = this.result_perplug12_divide/100;
        worksheet.getCell("L"+ty).numFmt = '0.00%'; 
      worksheet.getCell("L"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+ty).border = {
           top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      this.end_loss_plug12 = []
       for(var replug12_inch = 0; replug12_inch<resp.data.data.length; replug12_inch++){
      this.result_plug12_inch = (this.rowexport[replug12_inch].endless_length_yd*36)/(this.rowexport[replug12_inch].ttl_marker/this.rowexport[replug12_inch].length_ydb)
      
      this.end_loss_plug12.push(this.result_plug12_inch)
        
      }
         this.result_plug12_avg = 0;
        this.result_plug12_avg_sum = 0;
        
        for(var replug12_avg = 0; replug12_avg<resp.data.data.length; replug12_avg++){
        this.result_plug12_avg= this.end_loss_plug12[replug12_avg] * this.rowexport[replug12_avg].ttl_marker
        this.result_plug12_avg_sum = Number(this.result_plug12_avg_sum)+Number(this.result_plug12_avg)
        
      }  
     
        this.all_result_avg = this.result_plug12_avg_sum / this.result_ttlmarker
        this.keep_avg_loss_plug12.push({
          all_result_avg:this.all_result_avg
        })  

       /*   this.result_plug12 = 0
       this.result_plug12_sum = 0
      for(var re_plug1_2 = 0; re_plug1_2<resp.data.data.length; re_plug1_2++ ){
        this.result_plug12 = this.rowexport[re_plug1_2].plug12 * this.rowexport[re_plug1_2].ttl_marker
        this.result_plug12_sum = Number(this.result_plug12_sum)+Number(this.result_plug12)
      }
       this.result_plug12_divide = this.result_plug12_sum/this.result_ttlmarker
 */

        
      this.all_result_avg_last = parseFloat(this.all_result_avg).toFixed(2)
      worksheet.getCell("M"+ty).value = Number(this.all_result_avg_last);
      worksheet.getCell("M"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+ty).border = {
          top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      this.result_length_yd = 0
      for(var re_length_yd = 0; re_length_yd<resp.data.data.length; re_length_yd++ ){
        this.result_length_yd =  Number(this.result_length_yd) + Number(this.rowexport[re_length_yd].endless_length_yd)
        
      }
      this.keep_avg_length_yd.push({
       all_result_avg_length:this.result_length_yd
     })

     //this.result_length_yd = this.result_length_yd/resp.data.data.length 

      this.result_length_yd_last = parseFloat(this.result_length_yd).toFixed(2)
        worksheet.getCell("N"+ty).value = Number(this.result_length_yd_last);
      worksheet.getCell("N"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+ty).border = {
         top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };


      this.result_avg_per = 0;
      this.result_avg_per = this.result_length_yd / this.result_ttlmarker*100
      
      this.keep_avg_end.push({
       result_avg_per: this.result_avg_per
      })
      
      worksheet.getCell("O"+ty).value = this.result_avg_per/100;
      worksheet.getCell("O"+ty).numFmt = '0.00%'; 
      worksheet.getCell("O"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      this.result_splice = 0
      for(var re_splice = 0; re_splice<resp.data.data.length; re_splice++ ){
        this.result_splice =  Number(this.result_splice) + Number(this.rowexport[re_splice].spliceloss)
      }
     
    // this.result_splice = this.result_splice/resp.data.data.length
     this.keep_splice_Loss.push({
          result_splice:this.result_splice
     })

      this.result_splice = parseFloat(this.result_splice).toFixed(2)
        worksheet.getCell("P"+ty).value = Number(this.result_splice);
      worksheet.getCell("P"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      this.result_splicelossper = 0;
      this.result_splicelossper = this.result_splice / this.result_ttlmarker*100
       this.keep_splice_loss_per.push({
          result_splice_per:this.result_splicelossper
     })
      worksheet.getCell("Q"+ty).value = this.result_splicelossper/100
      worksheet.getCell("Q"+ty).numFmt = '0.00%'; 
      worksheet.getCell("Q"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };


       this.result_cut_out_loss = 0
      for(var re_cut_out = 0; re_cut_out<resp.data.data.length; re_cut_out++ ){
        this.result_cut_out_loss =  Number(this.result_cut_out_loss) + Number(this.rowexport[re_cut_out].totalcutout)
        
      }
     
     //this.result_cut_out_loss = this.result_cut_out_loss/resp.data.data.length 
     this.keep_cut_out_loss.push({
       result_cut_out_loss: this.result_cut_out_loss
     })

      this.result_cut_out_loss_last = parseFloat(this.result_cut_out_loss).toFixed(2)
        worksheet.getCell("R"+ty).value = Number(this.result_cut_out_loss_last);
      worksheet.getCell("R"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      
      this.result_cut_out_loss_per = 0;
      this.result_cut_out_loss_per = this.result_cut_out_loss / this.result_ttlmarker
      this.keep_cut_out_loss_per.push({
        result_cut_out_loss_per:this.result_cut_out_loss_per
      })
      worksheet.getCell("S"+ty).value = this.result_cut_out_loss_per/100
      worksheet.getCell("S"+ty).numFmt = '0.00%'; 
      worksheet.getCell("S"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      this.result_remnants = 0
      for(var re_remnants = 0; re_remnants<resp.data.data.length; re_remnants++ ){
        this.result_remnants =  Number(this.result_remnants) + Number(this.rowexport[re_remnants].rement)
        
      }
     
     //this.result_remnants = this.result_remnants/resp.data.data.length
     this.keep_remnants.push({
       result_remnants:this.result_remnants
     })

      this.result_remnants_last = parseFloat(this.result_remnants).toFixed(2)
        worksheet.getCell("T"+ty).value = Number(this.result_remnants_last);
      worksheet.getCell("T"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      this.result_remnants_per = 0;
      this.result_remnants_per = this.result_remnants / this.result_ttlmarker*100
       this.keep_remnants_per.push({
        remnants_per:this.result_remnants_per
      })
      
      worksheet.getCell("U"+ty).value = this.result_remnants_per/100;
      worksheet.getCell("U"+ty).numFmt = '0.00%'; 
      worksheet.getCell("U"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };


      this.result_total_length = 0
      for(var re_total_length = 0; re_total_length<resp.data.data.length; re_total_length++ ){
      this.result_total_length =  Number(this.result_total_length) + Number(this.rowexport[re_total_length].totallenthloss)
        
      }
     
    // this.result_total_length = this.result_total_length/resp.data.data.length
     this.keep_total_loss.push({
       result_total_length:this.result_total_length
     }) 

      this.result_total_length_last = parseFloat(this.result_total_length).toFixed(2)
        worksheet.getCell("V"+ty).value = Number(this.result_total_length_last);
      worksheet.getCell("V"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

       this.result_total_length_per = 0;
      this.result_total_length_per = this.result_total_length / this.result_ttlmarker*100
      this.keep_total_loss_per.push({
        result_total_loss_per:this.result_total_length_per
      })
      worksheet.getCell("W"+ty).value = this.result_total_length_per/100
      worksheet.getCell("W"+ty).numFmt = '0.00%'; 
      worksheet.getCell("W"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };




       //B---------------------------------------------------------------------

      worksheet.getCell("B"+a).value = "Table";
      worksheet.mergeCells("B"+a +":" + "B"+b);
      worksheet.getCell("B"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("B"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "B" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_table;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //C---------------------------------------------------------------------

       worksheet.getCell("C"+a).value = "S/O no.";
      worksheet.mergeCells("C"+a +":" + "C"+b);
      worksheet.getCell("C"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("C"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "C" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_no_short;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       
       //D---------------------------------------------------------------------

       worksheet.getCell("D"+a).value = "Customer";
      worksheet.mergeCells("D"+a +":" + "D"+b);
      worksheet.getCell("D"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("D"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "D" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].customer;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //E---------------------------------------------------------------------------
       
       worksheet.getCell("E"+a).value = "GM08 no.";
      worksheet.mergeCells("E"+a +":" + "E"+b);
      worksheet.getCell("E"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("E"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "E" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_no;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
       //F---------------------------------------------------------------
      if(this.multiple[x] == 1){
         this.multiple_rec = "Woven"
       }
       if(this.multiple[x] == 2){
         this.multiple_rec = "Light & Middle Weigth"
       }
       if(this.multiple[x] == 3){
         this.multiple_rec = "Fleece"
       }
       if(this.multiple[x] == 4){
         this.multiple_rec = "Special"
       }
    
      worksheet.getCell("F"+a).value = this.multiple_rec;
     worksheet.mergeCells("F"+a +":" + "F"+b);
      worksheet.getCell("F"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("F"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "F" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].fabric_type;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //G----------------------------------------------------------
       
       worksheet.getCell("G"+a).value = "Observe";
      worksheet.getCell("G"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("G"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("G"+b).value = "Width";
      worksheet.getCell("G"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("G"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "G" + ab
      
            var z = ab-a-2;
     this.last_value_obswidth = parseFloat(this.rowexport[z].obs_width).toFixed(2);
       worksheet.getCell(y).value = Number(this.last_value_obswidth)
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
  //H----------------------------------------------------------
        worksheet.getCell("H"+a).value = "Width";
      worksheet.getCell("H"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("H"+b).value = "(inch)";
      worksheet.getCell("H"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "H" + ab
      
            var z = ab-a-2;
      this.last_width_inch = parseFloat(this.rowexport[z].width_inch).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_width_inch);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //I--------------------------------------------------------------
         worksheet.getCell("I"+a).value = "Length";
      worksheet.getCell("I"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("I"+b).value = "(yd)";
      worksheet.getCell("I"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      
      this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "I" + ab
      
            var z = ab-a-2;
      this.last_length_ydb = parseFloat(this.rowexport[z].length_ydb).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_length_ydb);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
       //J-----------------------------------------------------------
         worksheet.getCell("J"+a).value = "Ttl Marker";
      worksheet.getCell("J"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("J"+b).value = "Length(yd)";
      worksheet.getCell("J"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "J" + ab
      
            var z = ab-a-2;
      this.last_ttl_marker = parseFloat(this.rowexport[z].ttl_marker).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_ttl_marker);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
       }; 
       }

        //K-------------------------------------------------------------------------------
         worksheet.getCell("K"+a).value = "Plug 1+2";
      worksheet.getCell("K"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("K"+b).value = "(inch)";
      worksheet.getCell("K"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
         
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "K" + ab
      
            var z = ab-a-2;
      this.last_plug12 = parseFloat(this.rowexport[z].plug12).toFixed(2)
      worksheet.getCell(y).value = Number(this.last_plug12);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
       }; 
       }

       //L------------------------------------------------------------------
            worksheet.getCell("L"+a).value = "%";
      worksheet.getCell("L"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("L"+b).value = 1.50/100;
       worksheet.getCell("L"+b).numFmt = '0.00%'; 
      worksheet.getCell("L"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

          this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "L" + ab
      
            var z = ab-a-2;
        this.last_widthloss = parseFloat(this.rowexport[z].widthloss).toFixed(2)
       worksheet.getCell(y).value = this.last_widthloss/100;
       worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
       }; 
       }
        //M--------------------------------------------------------------
        worksheet.getCell("M"+a).value = "Plug 1+2";
      worksheet.getCell("M"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("M"+b).value = "(inch)";
      worksheet.getCell("M"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "M" + ab
      
            var z = ab-a-2;

      this.result_plug12_inch = parseFloat((this.rowexport[z].endless_length_yd*36)/(this.rowexport[z].ttl_marker/this.rowexport[z].length_ydb)).toFixed(2)
      this.last_result_plug12_inch = parseFloat(this.result_plug12_inch).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_result_plug12_inch);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

 //N----------------------------------------------------------
       worksheet.getCell("N"+a).value = "Length";
      worksheet.getCell("N"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("N"+b).value = "(yd)";
      worksheet.getCell("N"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "N" + ab
      
            var z = ab-a-2;
      this.last_endless_yd = parseFloat(this.rowexport[z].endless_length_yd).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_endless_yd);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }


       //O------------------------------------------------------------------------
        worksheet.getCell("O"+a).value = "%";
      worksheet.getCell("O"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 

       if(this.multiple[x] == "Woven"){
        this.check_value = this.end_loss_value[0].p1
      }else if(this.multiple[x] == "Light-Middle Weight"){
        this.check_value = this.end_loss_value[0].p2
      }else if(this.multiple[x] == "Fleece"){
        this.check_value = this.end_loss_value[0].p3
      }else if(this.multiple[x] == "Special"){
        this.check_value = this.end_loss_value[0].p4
      }
      
      worksheet.getCell("O"+b).value = this.check_value/100
      worksheet.getCell("O"+b).numFmt = '0.00%'; 
      worksheet.getCell("O"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "O" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].avg_end/100;
       worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
      //P--------------------------------------------------------------------------
         worksheet.getCell("P"+a).value = "Length";
      worksheet.getCell("P"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("P"+b).value = "(yd)";
      worksheet.getCell("P"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "P" + ab
      
            var z = ab-a-2;
      this.last_splice_loss = parseFloat(this.rowexport[z].spliceloss).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_splice_loss);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //Q------------------------------------------------------------------------------
         worksheet.getCell("Q"+a).value = "%";
      worksheet.getCell("Q"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       if(this.multiple[x] == "Woven"){
        this.check_value = this.splice_value[0].p1
      }else if(this.multiple[x] == "Light-Middle Weight"){
        this.check_value = this.splice_value[0].p2
      }else if(this.multiple[x] == "Fleece"){
        this.check_value = this.splice_value[0].p3
      }else if(this.multiple[x] == "Special"){
        this.check_value = this.splice_value[0].p4
      }

       worksheet.getCell("Q"+b).value = this.check_value/100
       worksheet.getCell("Q"+b).numFmt = '0.00%'; 
      worksheet.getCell("Q"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "Q" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].splicelossper/100;
       worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //R------------------------------------------------------------------------------
         worksheet.getCell("R"+a).value = "Length";
      worksheet.getCell("R"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("R"+b).value = "(yd)";
      worksheet.getCell("R"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
          this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "R" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].totalcutout);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }


      //S------------------------------------------------------------------------------
         worksheet.getCell("S"+a).value = "%";
      worksheet.getCell("S"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
        if(this.multiple[x] == "Woven"){
        this.check_value = this.cut_out_value[0].p1
      }else if(this.multiple[x] == "Light-Middle Weight"){
        this.check_value = this.cut_out_value[0].p2
      }else if(this.multiple[x] == "Fleece"){
        this.check_value = this.cut_out_value[0].p3
      }else if(this.multiple[x] == "Special"){
        this.check_value = this.cut_out_value[0].p4
      }
       worksheet.getCell("S"+b).value = this.check_value/100;
       worksheet.getCell("S"+b).numFmt = '0.00%'; 
      worksheet.getCell("S"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      
        this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "S" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].totalcutoutper/100;
       worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //T-------------------------------------------------------------------------------

         worksheet.getCell("T"+a).value = "Length";
      worksheet.getCell("T"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("T"+b).value = "(yd)";
      worksheet.getCell("T"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
        this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "T" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].rement);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //U-------------------------------------------------------------------------------

         worksheet.getCell("U"+a).value = "%";
      worksheet.getCell("U"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
        if(this.multiple[x] == "Woven"){
        this.check_value = this.remnants_loss_value[0].p1
      }else if(this.multiple[x] == "Light-Middle Weight"){
        this.check_value = this.remnants_loss_value[0].p2
      }else if(this.multiple[x] == "Fleece"){
        this.check_value = this.remnants_loss_value[0].p3
      }else if(this.multiple[x] == "Special"){
        this.check_value = this.remnants_loss_value[0].p4
      } 
       worksheet.getCell("U"+b).value = this.check_value/100;
       worksheet.getCell("U"+b).numFmt = '0.00%'; 
      worksheet.getCell("U"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "U" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].rement_per/100;
      worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //V-------------------------------------------------------------------------------

         worksheet.getCell("V"+a).value = "Length";
      worksheet.getCell("V"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("V"+b).value = "(yd)";
      worksheet.getCell("V"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "V" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].totallenthloss);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }


       //W-------------------------------------------------------------------------------

         worksheet.getCell("W"+a).value = "%of";
      worksheet.getCell("W"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
         if(this.multiple[x] == "Woven"){
        this.check_value = this.total_length_value[0].p1
      }else if(this.multiple[x] == "Light-Middle Weight"){
        this.check_value = this.total_length_value[0].p2
      }else if(this.multiple[x] == "Fleece"){
        this.check_value = this.total_length_value[0].p3
      }else if(this.multiple[x] == "Special"){
        this.check_value = this.total_length_value[0].p4
      }
       worksheet.getCell("W"+b).value = this.check_value/100;
       worksheet.getCell("W"+b).numFmt = '0.00%'; 
      worksheet.getCell("W"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "W" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].totallenthlossper/100;
       worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

     

         } else{
          

             a = this.length+2
             b = this.length+3
     
      worksheet.getCell("A"+a).value = "Date";
      worksheet.mergeCells("A"+a +":" + "A"+b);
      worksheet.getCell("A"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "A" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].mu_date;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
      
       var ty = this.length+1
    if(this.multiple[x] == "1"){
        this.total_type = "Woven"
      }else if(this.multiple[x] == "2"){
        this.total_type = "Light-Middle Weight"
      }else if(this.multiple[x] == "3"){
        this.total_type = "Fleece"
      }else if(this.multiple[x] == "4"){
        this.total_type = "Special"
      }
      worksheet.getCell("A"+ty).value = "Total Type :"+this.total_type;
      worksheet.mergeCells("A"+ty +":" + "I"+ty);
      worksheet.getCell("A"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+ty).border = {
          top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
         this.result_ttlmarker = 0
      for(var re_ttlmarker = 0; re_ttlmarker<resp.data.data.length; re_ttlmarker++ ){
        this.result_ttlmarker =  Number(this.result_ttlmarker) + Number(this.rowexport[re_ttlmarker].ttl_marker)
        
      }
      this.keep_ttlmarker.push({
          result_ttlmarker:this.result_ttlmarker,
        })
        this.result_ttlmarker_2 = parseFloat(this.result_ttlmarker).toFixed(2)
        worksheet.getCell("J"+ty).value = Number(this.result_ttlmarker_2);
        worksheet.getCell("J"+ty).numFmt = "#,##0.00"
      worksheet.getCell("J"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+ty).border = {
          top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      this.result_plug12 = 0
       this.result_plug12_sum = 0
      for(var re_plug1_2 = 0; re_plug1_2<resp.data.data.length; re_plug1_2++ ){
        this.result_plug12 = this.rowexport[re_plug1_2].plug12 * this.rowexport[re_plug1_2].ttl_marker
        this.result_plug12_sum = Number(this.result_plug12_sum)+Number(this.result_plug12)
    
      }
       this.result_plug12_divide = this.result_plug12_sum/this.result_ttlmarker

    
     
     //this.result_plug12 = this.result_plug12/resp.data.data.length 
     this.keep_plug12_inch.push({
          result_plug12_inch:this.result_plug12_divide,
        })
        


        this.plug12_divide_last = parseFloat(this.result_plug12_divide).toFixed(2)
        worksheet.getCell("K"+ty).value = Number(this.plug12_divide_last);
      worksheet.getCell("K"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 

   this.result_perplug12 = 0
      this.result_perplug12_sum = 0

      for(var re_perplug12 = 0; re_perplug12<resp.data.data.length; re_perplug12++ ){
        this.result_perplug12 = this.rowexport[re_perplug12].widthloss * this.rowexport[re_perplug12].ttl_marker
        this.result_perplug12_sum =  Number(this.result_perplug12_sum) + Number(this.result_perplug12)
        
      }
      this.result_perplug12_divide =  this.result_perplug12_sum/this.result_ttlmarker
      this.keep_perplug12.push({
          result_perplug12:this.result_perplug12_divide,
        })

     //this.result_perplug12_l = this.result_perplug12/resp.data.data.length 
     

        worksheet.getCell("L"+ty).value = this.result_perplug12_divide/100;
        worksheet.getCell("L"+ty).numFmt = '0.00%'; 
      worksheet.getCell("L"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      this.end_loss_plug12 = []
     for(var replug12_inch = 0; replug12_inch<resp.data.data.length; replug12_inch++){
      this.result_plug12_inch = (this.rowexport[replug12_inch].endless_length_yd*36)/(this.rowexport[replug12_inch].ttl_marker/this.rowexport[replug12_inch].length_ydb)
    
      this.end_loss_plug12.push(this.result_plug12_inch)
        
      }
        this.result_plug12_avg = 0;
        this.result_plug12_avg_sum = 0;
        
        
        
        for(var replug12_avg = 0; replug12_avg<resp.data.data.length; replug12_avg++){
        this.result_plug12_avg= this.end_loss_plug12[replug12_avg] * this.rowexport[replug12_avg].ttl_marker
        this.result_plug12_avg_sum = Number(this.result_plug12_avg_sum)+Number(this.result_plug12_avg)
        
      }  
     
        this.all_result_avg = this.result_plug12_avg_sum / this.result_ttlmarker
        this.keep_avg_loss_plug12.push({
          all_result_avg:this.all_result_avg
        }) 

      this.all_result_avg_last = parseFloat(this.all_result_avg).toFixed(2)
      worksheet.getCell("M"+ty).value = Number(this.all_result_avg_last);
      worksheet.getCell("M"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      this.result_length_yd = 0
      for(var re_length_yd = 0; re_length_yd<resp.data.data.length; re_length_yd++ ){
        this.result_length_yd =  Number(this.result_length_yd) + Number(this.rowexport[re_length_yd].endless_length_yd)
        
      }
      this.keep_avg_length_yd.push({
       all_result_avg_length:this.result_length_yd
     })

     //this.result_length_yd = this.result_length_yd/resp.data.data.length 

      this.result_length_yd_last = parseFloat(this.result_length_yd).toFixed(2)
        worksheet.getCell("N"+ty).value = Number(this.result_length_yd_last);
      worksheet.getCell("N"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };


      
       this.result_avg_per = 0;
      this.result_avg_per = this.result_length_yd / this.result_ttlmarker*100
      
      this.keep_avg_end.push({
       result_avg_per: this.result_avg_per
      })
      
      worksheet.getCell("O"+ty).value = this.result_avg_per/100;
      worksheet.getCell("O"+ty).numFmt = '0.00%'; 
      worksheet.getCell("O"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      this.result_splice = 0
      for(var re_splice = 0; re_splice<resp.data.data.length; re_splice++ ){
        this.result_splice =  Number(this.result_splice) + Number(this.rowexport[re_splice].spliceloss)
        
      }
     
     //this.result_splice = this.result_splice/resp.data.data.length
     this.keep_splice_Loss.push({
          result_splice:this.result_splice
     })


      this.result_splice = parseFloat(this.result_splice).toFixed(2)
        worksheet.getCell("P"+ty).value = Number(this.result_splice);
      worksheet.getCell("P"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      this.result_splicelossper = 0;
      this.result_splicelossper = this.result_splice / this.result_ttlmarker*100
       this.keep_splice_loss_per.push({
          result_splice_per:this.result_splicelossper
     })
      worksheet.getCell("Q"+ty).value = this.result_splicelossper/100;
      worksheet.getCell("Q"+ty).numFmt = '0.00%'; 
      worksheet.getCell("Q"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };


       this.result_cut_out_loss = 0
      for(var re_cut_out = 0; re_cut_out<resp.data.data.length; re_cut_out++ ){
        this.result_cut_out_loss =  Number(this.result_cut_out_loss) + Number(this.rowexport[re_cut_out].totalcutout)
        
      }
     
     //this.result_cut_out_loss = this.result_cut_out_loss/resp.data.data.length 
     this.keep_cut_out_loss.push({
       result_cut_out_loss: this.result_cut_out_loss
     })
      
      this.result_cut_out_loss_last = parseFloat(this.result_cut_out_loss).toFixed(2)
        worksheet.getCell("R"+ty).value = Number(this.result_cut_out_loss_last);
      worksheet.getCell("R"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      
      this.result_cut_out_loss_per = 0;
      this.result_cut_out_loss_per = this.result_cut_out_loss / this.result_ttlmarker
      this.keep_cut_out_loss_per.push({
        result_cut_out_loss_per:this.result_cut_out_loss_per
      })
      worksheet.getCell("S"+ty).value = this.result_cut_out_loss_per/100;
      worksheet.getCell("S"+ty).numFmt = '0.00%'; 
      worksheet.getCell("S"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      this.result_remnants = 0
      for(var re_remnants = 0; re_remnants<resp.data.data.length; re_remnants++ ){
        this.result_remnants =  Number(this.result_remnants) + Number(this.rowexport[re_remnants].rement)
        
      }
     
     //this.result_remnants = this.result_remnants/resp.data.data.length
     this.keep_remnants.push({
       result_remnants:this.result_remnants
     })

      this.result_remnants_last = parseFloat(this.result_remnants).toFixed(2)
        worksheet.getCell("T"+ty).value = Number(this.result_remnants_last);
      worksheet.getCell("T"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      this.result_remnants_per = 0;
      this.result_remnants_per = this.result_remnants / this.result_ttlmarker*100
       this.keep_remnants_per.push({
        remnants_per:this.result_remnants_per
      })
      worksheet.getCell("U"+ty).value = this.result_remnants_per;
      worksheet.getCell("U"+ty).numFmt = '0.00%'; 
      worksheet.getCell("U"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };


      this.result_total_length = 0
      for(var re_total_length = 0; re_total_length<resp.data.data.length; re_total_length++ ){
      this.result_total_length =  Number(this.result_total_length) + Number(this.rowexport[re_total_length].totallenthloss)
        
      }
     
    // this.result_total_length = this.result_total_length/resp.data.data.length
     this.keep_total_loss.push({
       result_total_length:this.result_total_length
     }) 

      this.result_total_length_last = parseFloat(this.result_total_length).toFixed(2)
        worksheet.getCell("V"+ty).value = Number(this.result_total_length_last);
      worksheet.getCell("V"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

       this.result_total_length_per = 0;
      this.result_total_length_per = this.result_total_length / this.result_ttlmarker*100
      this.keep_total_loss_per.push({
        result_total_loss_per:this.result_total_length_per
      })
      worksheet.getCell("W"+ty).value = this.result_total_length_per/100;
      worksheet.getCell("W"+ty).numFmt = '0.00%'; 
      worksheet.getCell("W"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+ty).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };




       //B----------------------------------------------------------------
      worksheet.getCell("B"+a).value = "Table";
      worksheet.mergeCells("B"+a +":" + "B"+b);
      worksheet.getCell("B"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("B"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "B" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_table;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //C---------------------------------------------------------------------

       worksheet.getCell("C"+a).value = "S/O no.";
      worksheet.mergeCells("C"+a +":" + "C"+b);
      worksheet.getCell("C"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("C"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "C" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_no_short;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

           //D---------------------------------------------------------------------

       worksheet.getCell("D"+a).value = "Customer";
      worksheet.mergeCells("D"+a +":" + "D"+b);
      worksheet.getCell("D"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("D"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "D" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].customer;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

         //E---------------------------------------------------------------------------
       
       worksheet.getCell("E"+a).value = "GM08 no.";
      worksheet.mergeCells("E"+a +":" + "E"+b);
      worksheet.getCell("E"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("E"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "E" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_no;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //F---------------------------------------------------------------
     if(this.multiple[x] == 1){
         this.multiple_rec = "Woven"
       }
       if(this.multiple[x] == 2){
         this.multiple_rec = "Light & Middle Weigth"
       }
       if(this.multiple[x] == 3){
         this.multiple_rec = "Fleece"
       }
       if(this.multiple[x] == 4){
         this.multiple_rec = "Special"
       }
    
      worksheet.getCell("F"+a).value = this.multiple_rec;
      worksheet.mergeCells("F"+a +":" + "F"+b);
      worksheet.getCell("F"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("F"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "F" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].fabric_type;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //G----------------------------------------------------------
       
       worksheet.getCell("G"+a).value = "Observe";
      worksheet.getCell("G"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("G"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("G"+b).value = "Width";
      worksheet.getCell("G"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("G"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "G" + ab
      
            var z = ab-a-2;
      this.last_value_obswidth = parseFloat(this.rowexport[z].obs_width).toFixed(2);
       worksheet.getCell(y).value = Number(this.last_value_obswidth)
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //H----------------------------------------------------------
        worksheet.getCell("H"+a).value = "Width";
      worksheet.getCell("H"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("H"+b).value = "(inch)";
      worksheet.getCell("H"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "H" + ab
      
            var z = ab-a-2;
      this.last_width_inch = parseFloat(this.rowexport[z].width_inch).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_width_inch);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
       //I--------------------------------------------------------------
         worksheet.getCell("I"+a).value = "Length";
      worksheet.getCell("I"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("I"+b).value = "(yd)";
      worksheet.getCell("I"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      
      this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "I" + ab
      
            var z = ab-a-2;
      this.last_length_ydb = parseFloat(this.rowexport[z].length_ydb).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_length_ydb);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

         //J-----------------------------------------------------------
         worksheet.getCell("J"+a).value = "Ttl Marker";
      worksheet.getCell("J"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("J"+b).value = "Length(yd)";
      worksheet.getCell("J"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "J" + ab
      
            var z = ab-a-2;
     this.last_ttl_marker = parseFloat(this.rowexport[z].ttl_marker).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_ttl_marker);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
       }; 
       }

        //K-------------------------------------------------------------------------------
         worksheet.getCell("K"+a).value = "Plug 1+2";
      worksheet.getCell("K"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("K"+b).value = "(inch)";
      worksheet.getCell("K"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
         
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "K" + ab
      
            var z = ab-a-2;
      this.last_plug12 = parseFloat(this.rowexport[z].plug12).toFixed(2)
      worksheet.getCell(y).value = Number(this.last_plug12);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
       }; 
       }

       //L------------------------------------------------------------------
            worksheet.getCell("L"+a).value = "%";
      worksheet.getCell("L"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("L"+b).value = 1.50/100;
       worksheet.getCell("L"+b).numFmt = '0.00%'; 
      worksheet.getCell("L"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

          this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "L" + ab
      
            var z = ab-a-2;
        this.last_widthloss = parseFloat(this.rowexport[z].widthloss).toFixed(2)
      worksheet.getCell(y).value = this.last_widthloss/100;
      worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
       }; 
       }

        //M--------------------------------------------------------------
        worksheet.getCell("M"+a).value = "Plug 1+2";
      worksheet.getCell("M"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("M"+b).value = "(inch)";
      worksheet.getCell("M"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "M" + ab
      
            var z = ab-a-2;

      this.result_plug12_inch = parseFloat((this.rowexport[z].endless_length_yd*36)/(this.rowexport[z].ttl_marker/this.rowexport[z].length_ydb)).toFixed(2)
      this.last_result_plug12_inch = parseFloat(this.result_plug12_inch).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_result_plug12_inch);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }


      //N----------------------------------------------------------
       worksheet.getCell("N"+a).value = "Length";
      worksheet.getCell("N"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("N"+b).value = "(yd)";
      worksheet.getCell("N"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "N" + ab
      
            var z = ab-a-2;

      this.last_endless_yd = parseFloat(this.rowexport[z].endless_length_yd).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_endless_yd);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }


       //O------------------------------------------------------------------------
        worksheet.getCell("O"+a).value = "%";
      worksheet.getCell("O"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 

       if(this.multiple[x] == "Woven"){
        this.check_value = this.end_loss_value[0].p1
      }else if(this.multiple[x] == "Light-Middle Weight"){
        this.check_value = this.end_loss_value[0].p2
      }else if(this.multiple[x] == "Fleece"){
        this.check_value = this.end_loss_value[0].p3
      }else if(this.multiple[x] == "Special"){
        this.check_value = this.end_loss_value[0].p4
      }
      
      worksheet.getCell("O"+b).value = this.check_value/100;
      worksheet.getCell("O"+b).numFmt = '0.00%'; 
      worksheet.getCell("O"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "O" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].avg_end/100;
       worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //P--------------------------------------------------------------------------
         worksheet.getCell("P"+a).value = "Length";
      worksheet.getCell("P"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("P"+b).value = "(yd)";
      worksheet.getCell("P"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "P" + ab
      
            var z = ab-a-2;
      this.last_splice_loss = parseFloat(this.rowexport[z].spliceloss).toFixed(2)
       worksheet.getCell(y).value = Number(this.last_splice_loss);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //Q------------------------------------------------------------------------------
         worksheet.getCell("Q"+a).value = "%";
      worksheet.getCell("Q"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       if(this.multiple[x] == "Woven"){
        this.check_value = this.splice_value[0].p1
      }else if(this.multiple[x] == "Light-Middle Weight"){
        this.check_value = this.splice_value[0].p2
      }else if(this.multiple[x] == "Fleece"){
        this.check_value = this.splice_value[0].p3
      }else if(this.multiple[x] == "Special"){
        this.check_value = this.splice_value[0].p4
      }

       worksheet.getCell("Q"+b).value = this.check_value/100;
       worksheet.getCell("Q"+b).numFmt = '0.00%'; 
      worksheet.getCell("Q"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "Q" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].splicelossper/100;
       worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //R------------------------------------------------------------------------------
         worksheet.getCell("R"+a).value = "Length";
      worksheet.getCell("R"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("R"+b).value = "(yd)";
      worksheet.getCell("R"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
         this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "R" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].totalcutout);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //S------------------------------------------------------------------------------
         worksheet.getCell("S"+a).value = "%";
      worksheet.getCell("S"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       if(this.multiple[x] == "Woven"){
        this.check_value = this.cut_out_value[0].p1
      }else if(this.multiple[x] == "Light-Middle Weight"){
        this.check_value = this.cut_out_value[0].p2
      }else if(this.multiple[x] == "Fleece"){
        this.check_value = this.cut_out_value[0].p3
      }else if(this.multiple[x] == "Special"){
        this.check_value = this.cut_out_value[0].p4
      }

       worksheet.getCell("S"+b).value = this.check_value/100;
       worksheet.getCell("S"+b).numFmt = '0.00%'; 
      worksheet.getCell("S"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

        this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "S" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].totalcutoutper/100;
       worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //T-------------------------------------------------------------------------------

         worksheet.getCell("T"+a).value = "Length";
      worksheet.getCell("T"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("T"+b).value = "(yd)";
      worksheet.getCell("T"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "T" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].rement);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //U-------------------------------------------------------------------------------

         worksheet.getCell("U"+a).value = "%";
      worksheet.getCell("U"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
        if(this.multiple[x] == "Woven"){
        this.check_value = this.remnants_loss_value[0].p1
      }else if(this.multiple[x] == "Light-Middle Weight"){
        this.check_value = this.remnants_loss_value[0].p2
      }else if(this.multiple[x] == "Fleece"){
        this.check_value = this.remnants_loss_value[0].p3
      }else if(this.multiple[x] == "Special"){
        this.check_value = this.remnants_loss_value[0].p4
      }
       worksheet.getCell("U"+b).value = this.check_value/100;
       worksheet.getCell("U"+b).numFmt = '0.00%'; 
      worksheet.getCell("U"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "U" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].rement_per/100;
       worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //V-------------------------------------------------------------------------------

         worksheet.getCell("V"+a).value = "Length";
      worksheet.getCell("V"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
       worksheet.getCell("V"+b).value = "(yd)";
      worksheet.getCell("V"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+b).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "V" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].totallenthloss);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //W-------------------------------------------------------------------------------

         worksheet.getCell("W"+a).value = "%of";
      worksheet.getCell("W"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+a).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      }; 
         if(this.multiple[x] == "Woven"){
        this.check_value = this.total_length_value[0].p1
      }else if(this.multiple[x] == "Light-Middle Weight"){
        this.check_value = this.total_length_value[0].p2
      }else if(this.multiple[x] == "Fleece"){
        this.check_value = this.total_length_value[0].p3
      }else if(this.multiple[x] == "Special"){
        this.check_value = this.total_length_value[0].p4
      }
       worksheet.getCell("W"+b).value = this.check_value/100;
       worksheet.getCell("W"+b).numFmt = '0.00%'; 
      worksheet.getCell("W"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "W" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].totallenthlossper/100;
       worksheet.getCell(y).numFmt = '0.00%'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      
         }  
         

      
         count++
          }
       
         if(this.multiple.length == count){
          
      var c = this.length+2
      var d = this.length+3     
      var e = this.length+5
      var f = this.length+6
      var g = this.length+7
      var h = this.length+8
      var i = this.length+9
      var j = this.length+10
      var k = this.length+11
      var l = this.length+12 
      var m = this.length+13 
      var n = this.length+14 
      var o = this.length+15 //endloss
      var p = this.length+16 //splice loss
      var q = this.length+17 //Cut - out Loss
      var r = this.length+18 //Remnants Loss
      var s = this.length+19 //Total Loss
        for(var xyz1 = 2; xyz1<s+1; xyz1++){
        worksheet.getRow(xyz1).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 11,
        bold: false
      };
        }

      worksheet.getCell("A"+c).value = "Total Target Weight";
      worksheet.mergeCells("A"+c+":" +"I"+c);
      worksheet.getCell("A"+c).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+c).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

        worksheet.getCell("J"+c).value = "";
        worksheet.getCell("J"+c).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+c).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

        worksheet.getCell("K"+c).value = "";
        worksheet.getCell("K"+c).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+c).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

        worksheet.getCell("L"+c).value = 1.50/100;
        worksheet.getCell("L"+c).numFmt = '0.00%'; 
        worksheet.getCell("L"+c).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+c).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

        worksheet.getCell("M"+c).value = "";
        worksheet.getCell("M"+c).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+c).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
        worksheet.getCell("N"+c).value = "";
        worksheet.getCell("N"+c).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+c).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

     
      this.per = 0
      
      for(var ak7 = 0; ak7<this.keep_multi.length; ak7++){
        
         if(this.keep_multi[ak7].multiple == "Woven"){
           
           if(this.keep_ttlmarker.length == 4){
          this.per = (this.end_loss_value[ak7].p1)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_percent.push({
            total_percent:this.per
          })
        } else if(this.keep_ttlmarker.length == 3){
          this.per = (this.end_loss_value[ak7].p1)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_percent.push({
            total_percent:this.per
          })
        }
        else if(this.keep_ttlmarker.length == 2){
          this.per = (this.end_loss_value[ak7].p1)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_percent.push({
            total_percent:this.per
          })
        }
        else if(this.keep_ttlmarker.length == 1){
          this.per = (this.end_loss_value[ak7].p1)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_percent.push({
            total_percent:this.per
          })
        } 
        }else if(this.keep_multi[ak7].multiple == "Light-Middle Weight"){
          if(this.keep_ttlmarker.legnth == 4){
          this.per = (this.end_loss_value[0].p2)*(this.keep_ttlmarker[1].result_ttlmarker)
        
          this.keep_total_percent.push({
            total_percent:this.per
          })
          }else if(this.keep_ttlmarker.length == 3){
             this.per = (this.end_loss_value[0].p2)*(this.keep_ttlmarker[0].result_ttlmarker)
         
          this.keep_total_percent.push({
            total_percent:this.per
          })
          }
          else if(this.keep_ttlmarker.length == 2){
             this.per = (this.end_loss_value[0].p2)*(this.keep_ttlmarker[0].result_ttlmarker)
         
          this.keep_total_percent.push({
            total_percent:this.per
          })
          }
          else if(this.keep_ttlmarker.length == 1){
             this.per = (this.end_loss_value[0].p2)*(this.keep_ttlmarker[0].result_ttlmarker)
         
          this.keep_total_percent.push({
            total_percent:this.per
          })
          }
        }else if(this.keep_multi[ak7].multiple == "Fleece"){
         if(this.keep_ttlmarker.legnth == 4){
          this.per = (this.end_loss_value[0].p3)*(this.keep_ttlmarker[2].result_ttlmarker)
     
          this.keep_total_percent.push({
            total_percent:this.per
          })
          }else if(this.keep_ttlmarker.length == 3){
             this.per = (this.end_loss_value[0].p3)*(this.keep_ttlmarker[1].result_ttlmarker)

          this.keep_total_percent.push({
            total_percent:this.per
          })
          }
          else if(this.keep_ttlmarker.length == 2){
             this.per = (this.end_loss_value[0].p3)*(this.keep_ttlmarker[0].result_ttlmarker)
          
          this.keep_total_percent.push({
            total_percent:this.per
          })
          }
          else if(this.keep_ttlmarker.length == 1){
             this.per = (this.end_loss_value[0].p3)*(this.keep_ttlmarker[0].result_ttlmarker)
         
          this.keep_total_percent.push({
            total_percent:this.per
          })
          }
        }else if(this.keep_multi[ak7].multiple == "Special"){
         if(this.keep_ttlmarker.legnth == 4){
          this.per = (this.end_loss_value[0].p4)*(this.keep_ttlmarker[3].result_ttlmarker)
        
          this.keep_total_percent.push({
            total_percent:this.per
          })
          }else if(this.keep_ttlmarker.length == 3){
             this.per = (this.end_loss_value[0].p4)*(this.keep_ttlmarker[2].result_ttlmarker)
          
          this.keep_total_percent.push({
            total_percent:this.per
          })
          }
          else if(this.keep_ttlmarker.length == 2){
             this.per = (this.end_loss_value[0].p4)*(this.keep_ttlmarker[1].result_ttlmarker)
        
          this.keep_total_percent.push({
            total_percent:this.per
          })
          }
          else if(this.keep_ttlmarker.length == 1){
             this.per = (this.end_loss_value[0].p4)*(this.keep_ttlmarker[0].result_ttlmarker)
        
          this.keep_total_percent.push({
            total_percent:this.per
          })
          }
        } 
      }
       this.result_sum_ttlmarker = 0
      for(var jd = 0; jd<this.keep_ttlmarker.length; jd++){
        this.result_sum_ttlmarker = Number(this.result_sum_ttlmarker)+Number(this.keep_ttlmarker[jd].result_ttlmarker)
       
      }
      
      this.sum_keep_total_percent = 0
      for(var axb = 0; axb<this.keep_total_percent.length; axb++){
        this.sum_keep_total_percent = Number(this.sum_keep_total_percent)+this.keep_total_percent[axb].total_percent
      }
      this.last_keep_total_percent = 0
      
      this.last_keep_total_percent = this.sum_keep_total_percent/this.result_sum_ttlmarker;


             this.result_overall_end_loss = 0
   
    this.result_overall_end_loss =  Number(this.end_loss_value[0].p1*this.total_overall[0].result_ttl)
    +Number(this.end_loss_value[0].p2*this.total_overall[1].result_ttl)
    +Number(this.end_loss_value[0].p3*this.total_overall[2].result_ttl)
    +Number(this.end_loss_value[0].p4*this.total_overall[3].result_ttl)
    
    this.result_overall_end_loss = this.result_overall_end_loss / this.result_sum_ttlmarker

    this.all_value_before_target.push({
      value:this.result_overall_end_loss,
      value_name:"End Loss",
      value_target:this.result_overall_end_loss
    })

       worksheet.getCell("O"+c).value = this.result_overall_end_loss/100;
       worksheet.getCell("O"+c).numFmt = '0.00%'; 
       worksheet.getCell("O"+c).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+c).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

        worksheet.getCell("P"+c).value = "";
        worksheet.getCell("P"+c).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+c).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

       this.result_overall_splice_loss = 0
    this.result_overall_splice_loss =  Number(this.splice_value[0].p1*this.total_overall[0].result_ttl)
    +Number(this.splice_value[0].p2*this.total_overall[1].result_ttl)
    +Number(this.splice_value[0].p3*this.total_overall[2].result_ttl)
    +Number(this.splice_value[0].p4*this.total_overall[3].result_ttl)

    
    
    this.result_overall_splice_loss = this.result_overall_splice_loss / this.result_sum_ttlmarker
   this.all_value_before_target.push({
      value:this.result_overall_splice_loss,
      value_name:"Splice Loss",
      value_target:this.result_overall_splice_loss
    })

      worksheet.getCell("Q"+c).value = this.result_overall_splice_loss/100;
      worksheet.getCell("Q"+c).numFmt = '0.00%'; 
      worksheet.getCell("Q"+c).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+c).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

        worksheet.getCell("R"+c).value = "";
        worksheet.getCell("R"+c).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+c).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

    this.result_overall_cut = 0
    this.result_overall_cut =  Number(this.cut_out_value[0].p1*this.total_overall[0].result_ttl)
    +Number(this.cut_out_value[0].p2*this.total_overall[1].result_ttl)
    +Number(this.cut_out_value[0].p3*this.total_overall[2].result_ttl)
    +Number(this.cut_out_value[0].p4*this.total_overall[3].result_ttl)
    this.result_overall_cut = this.result_overall_cut / this.result_sum_ttlmarker
      this.all_value_before_target.push({
      value:this.result_overall_cut,
      value_name:"Cut Out Loss",
      value_target:this.result_overall_cut
    })
    
      worksheet.getCell("S"+c).value = this.result_overall_cut/100;
      worksheet.getCell("S"+c).numFmt = '0.00%'; 
      worksheet.getCell("S"+c).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+c).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
        worksheet.getCell("T"+c).value = "";
        worksheet.getCell("T"+c).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+c).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

    /*    remnants_loss_value:[
        {
                p1:0.05,
                p2:0.08,
                p3:0.11,
                p4:0.10,
        }
        ], */
      
      this.per2 = 0
      for(var ak7 = 0; ak7<this.keep_multi.length; ak7++){
        
         if(this.keep_multi[ak7].multiple == "Woven"){
           if(this.keep_ttlmarker.length == 4){
          this.per2 = (this.remnants_loss_value[0].p1)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_rem_percent.push({
            total_percent:this.per2
          })
        } else if(this.keep_ttlmarker.length == 3){
          this.per2 = (this.remnants_loss_value[0].p1)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_rem_percent.push({
            total_percent:this.per2
          })
        }
        else if(this.keep_ttlmarker.length == 2){
          this.per2 = (this.remnants_loss_value[0].p1)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_rem_percent.push({
            total_percent:this.per2
          })
        }
        else if(this.keep_ttlmarker.length == 1){
          this.per2 = (this.remnants_loss_value[0].p1)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_rem_percent.push({
            total_percent:this.per2
          })
        } 
        }else if(this.keep_multi[ak7].multiple == "Light-Middle Weight"){
          if(this.keep_ttlmarker.legnth == 4){
          this.per2 = (this.remnants_loss_value[0].p2)*(this.keep_ttlmarker[1].result_ttlmarker)
        
          this.keep_total_rem_percent.push({
            total_percent:this.per2
          })
          }else if(this.keep_ttlmarker.length == 3){
           this.per2 = (this.remnants_loss_value[0].p2)*(this.keep_ttlmarker[0].result_ttlmarker)
        
          this.keep_total_rem_percent.push({
            total_percent:this.per2
          })
          }
          else if(this.keep_ttlmarker.length == 2){
         this.per2 = (this.remnants_loss_value[0].p2)*(this.keep_ttlmarker[0].result_ttlmarker)
        
          this.keep_total_rem_percent.push({
            total_percent:this.per2
          })
          }
          else if(this.keep_ttlmarker.length == 1){
            this.per2 = (this.remnants_loss_value[0].p2)*(this.keep_ttlmarker[0].result_ttlmarker)
        
          this.keep_total_rem_percent.push({
            total_percent:this.per2
          })
          }
        }else if(this.keep_multi[ak7].multiple == "Fleece"){
         if(this.keep_ttlmarker.legnth == 4){
          this.per2 = (this.remnants_loss_value[0].p3)*(this.keep_ttlmarker[2].result_ttlmarker)
     
          this.keep_total_rem_percent.push({
            total_percent:this.per2
          })
          }else if(this.keep_ttlmarker.length == 3){
          this.per2 = (this.remnants_loss_value[0].p3)*(this.keep_ttlmarker[1].result_ttlmarker)
     
          this.keep_total_rem_percent.push({
            total_percent:this.per2
          })
          }
          else if(this.keep_ttlmarker.length == 2){
          this.per2 = (this.remnants_loss_value[0].p3)*(this.keep_ttlmarker[0].result_ttlmarker)
     
          this.keep_total_rem_percent.push({
            total_percent:this.per2
          })
          }
          else if(this.keep_ttlmarker.length == 1){
          this.per2 = (this.remnants_loss_value[0].p3)*(this.keep_ttlmarker[0].result_ttlmarker)
     
          this.keep_total_rem_percent.push({
            total_percent:this.per2
          })
          }
        }else if(this.keep_multi[ak7].multiple == "Special"){
         if(this.keep_ttlmarker.legnth == 4){
          this.per2 = (this.remnants_loss_value[0].p4)*(this.keep_ttlmarker[3].result_ttlmarker)
        
          this.keep_total_rem_percent.push({
            total_percent:this.per2
          })
          }else if(this.keep_ttlmarker.length == 3){
           this.per2 = (this.remnants_loss_value[0].p4)*(this.keep_ttlmarker[2].result_ttlmarker)
        
          this.keep_total_rem_percent.push({
            total_percent:this.per2
          })
          }
          else if(this.keep_ttlmarker.length == 2){
            this.per2 = (this.remnants_loss_value[0].p4)*(this.keep_ttlmarker[1].result_ttlmarker)
        
          this.keep_total_rem_percent.push({
            total_percent:this.per2
          })
          }
          else if(this.keep_ttlmarker.length == 1){
            this.per2 = (this.remnants_loss_value[0].p4)*(this.keep_ttlmarker[0].result_ttlmarker)
        
          this.keep_total_rem_percent.push({
            total_percent:this.per2
          })
          }
        } 
      }
       this.result_sum_ttlmarker = 0
      for(var jd = 0; jd<this.keep_ttlmarker.length; jd++){
        this.result_sum_ttlmarker = Number(this.result_sum_ttlmarker)+Number(this.keep_ttlmarker[jd].result_ttlmarker)
       
      }
      
      this.sum_keep_total_rem_percent = 0
      for(var axb = 0; axb<this.keep_total_rem_percent.length; axb++){
        this.sum_keep_total_rem_percent = Number(this.sum_keep_total_rem_percent)+Number(this.keep_total_rem_percent[axb].total_percent)
      }
      this.last_keep_total_rem_percent = 0
      
      this.last_keep_total_rem_percent = this.sum_keep_total_rem_percent/this.result_sum_ttlmarker;

      this.result_overall_rem = 0
    this.result_overall_rem =  Number(this.remnants_loss_value[0].p1*this.total_overall[0].result_ttl)
    +Number(this.remnants_loss_value[0].p2*this.total_overall[1].result_ttl)
    +Number(this.remnants_loss_value[0].p3*this.total_overall[2].result_ttl)
    +Number(this.remnants_loss_value[0].p4*this.total_overall[3].result_ttl)
    this.result_overall_rem = this.result_overall_rem / this.result_sum_ttlmarker

    this.all_value_before_target.push({
      value:this.result_overall_rem,
      value_name:"Remnants Loss",
      value_target:this.result_overall_rem
    })
    
    

      worksheet.getCell("U"+c).value = this.result_overall_rem/100;
      worksheet.getCell("U"+c).numFmt = '0.00%'; 
      worksheet.getCell("U"+c).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+c).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
        worksheet.getCell("V"+c).value = "";
        worksheet.getCell("V"+c).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+c).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

/*        keep_total_loss_per:[],
        keep_total_total_length_percent:[], */
      /*   total_length_value:[
          {
                p1:0.38,
                p2:0.50,
                p3:0.56,
                p4:0.65,
        }
        ] */

            this.per3 = 0
      for(var ak7 = 0; ak7<this.keep_multi.length; ak7++){
        
         if(this.keep_multi[ak7].multiple == "Woven"){
           if(this.keep_ttlmarker.length == 4){
          this.per3 = (this.total_length_value[0].p1)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_total_length_percent.push({
            total_percent:this.per3
          })
        } else if(this.keep_ttlmarker.length == 3){
           this.per3 = (this.total_length_value[0].p1)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_total_length_percent.push({
            total_percent:this.per3
          })
        }else if(this.keep_ttlmarker.length == 2){
          this.per3 = (this.total_length_value[0].p1)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_total_length_percent.push({
            total_percent:this.per3
          })
        }
        else if(this.keep_ttlmarker.length == 1){
           this.per3 = (this.total_length_value[0].p1)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_total_length_percent.push({
            total_percent:this.per3
          })
        } 
        }else if(this.keep_multi[ak7].multiple == "Light-Middle Weight"){
          if(this.keep_ttlmarker.legnth == 4){
           this.per3 = (this.total_length_value[0].p2)*(this.keep_ttlmarker[1].result_ttlmarker)
          this.keep_total_total_length_percent.push({
            total_percent:this.per3
          })
          }else if(this.keep_ttlmarker.length == 3){
           this.per3 = (this.total_length_value[0].p2)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_total_length_percent.push({
            total_percent:this.per3
          })
          }
          else if(this.keep_ttlmarker.length == 2){
          this.per3 = (this.total_length_value[0].p2)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_total_length_percent.push({
            total_percent:this.per3
          })
          }
          else if(this.keep_ttlmarker.length == 1){
          this.per3 = (this.total_length_value[0].p2)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_total_length_percent.push({
            total_percent:this.per3
          })
          }
        }else if(this.keep_multi[ak7].multiple == "Fleece"){
         if(this.keep_ttlmarker.legnth == 4){
           this.per3 = (this.total_length_value[0].p3)*(this.keep_ttlmarker[2].result_ttlmarker)
          this.keep_total_total_length_percent.push({
            total_percent:this.per3
          })
          }else if(this.keep_ttlmarker.length == 3){
           this.per3 = (this.total_length_value[0].p3)*(this.keep_ttlmarker[1].result_ttlmarker)
          this.keep_total_total_length_percent.push({
            total_percent:this.per3
          })
          }
          else if(this.keep_ttlmarker.length == 2){
           this.per3 = (this.total_length_value[0].p3)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_total_length_percent.push({
            total_percent:this.per3
          })
          }
          else if(this.keep_ttlmarker.length == 1){
           this.per3 = (this.total_length_value[0].p3)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_total_length_percent.push({
            total_percent:this.per3
          })
          }
        }else if(this.keep_multi[ak7].multiple == "Special"){
         if(this.keep_ttlmarker.legnth == 4){
           this.per3 = (this.total_length_value[0].p4)*(this.keep_ttlmarker[3].result_ttlmarker)
          this.keep_total_total_length_percent.push({
            total_percent:this.per3
          })
          }else if(this.keep_ttlmarker.length == 3){
          this.per3 = (this.total_length_value[0].p4)*(this.keep_ttlmarker[2].result_ttlmarker)
          this.keep_total_total_length_percent.push({
            total_percent:this.per3
          })
          }
          else if(this.keep_ttlmarker.length == 2){
           this.per3 = (this.total_length_value[0].p4)*(this.keep_ttlmarker[1].result_ttlmarker)
          this.keep_total_total_length_percent.push({
            total_percent:this.per3
          })
          }
          else if(this.keep_ttlmarker.length == 1){
          this.per3 = (this.total_length_value[0].p4)*(this.keep_ttlmarker[0].result_ttlmarker)
          this.keep_total_total_length_percent.push({
            total_percent:this.per3
          })
          }
        } 
      }
       this.result_sum_ttlmarker = 0
      for(var jd = 0; jd<this.keep_ttlmarker.length; jd++){
        this.result_sum_ttlmarker = Number(this.result_sum_ttlmarker)+Number(this.keep_ttlmarker[jd].result_ttlmarker)
       
      }
     
      this.sum_keep_total_total_percent = 0
      for(var axb = 0; axb<this.keep_total_total_length_percent.length; axb++){
        this.sum_keep_total_total_percent = Number(this.sum_keep_total_total_percent)+Number(this.keep_total_total_length_percent[axb].total_percent)
      }
      this.last_keep_total_total_percent = 0
  
      this.last_keep_total_total_percent = this.sum_keep_total_total_percent/this.result_sum_ttlmarker;

       this.result_overall_total = 0
    
     this.result_overall_total =  Number(this.total_length_value[0].p1*this.total_overall[0].result_ttl)
    +Number(this.total_length_value[0].p2*this.total_overall[1].result_ttl)
    +Number(this.total_length_value[0].p3*this.total_overall[2].result_ttl)
    +Number(this.total_length_value[0].p4*this.total_overall[3].result_ttl)
    
    this.result_overall_total = this.result_overall_total / this.result_sum_ttlmarker

      worksheet.getCell("W"+c).value = this.result_overall_total/100;
      worksheet.getCell("W"+c).numFmt = '0.00%'; 
      worksheet.getCell("W"+c).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+c).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      let org = this.$q.localStorage.getItem("org");
     
      worksheet.getCell("A"+d).value = "NY"+org+ " : Daily Total";
      worksheet.mergeCells("A"+d+":" +"I"+d);
      worksheet.getCell("A"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+d).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      this.result_sum_ttlmarker = 0
      for(var jd = 0; jd<this.keep_ttlmarker.length; jd++){
        this.result_sum_ttlmarker = Number(this.result_sum_ttlmarker)+Number(this.keep_ttlmarker[jd].result_ttlmarker)
       
      }
      this.last_result_sum_ttlmarker = parseFloat(this.result_sum_ttlmarker).toFixed(2)
      
      worksheet.getCell("J"+d).value = Number(this.last_result_sum_ttlmarker);
      worksheet.getCell("J"+d).numFmt = "#,##0.00"
      worksheet.getCell("J"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+d).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
   

        this.sum_plug12_inch = 0
        this.result_sum_plug12_inch = 0
       
      for(var kd = 0; kd<this.keep_ttlmarker.length; kd++){
         this.result_sum_plug12_inch = this.keep_ttlmarker[kd].result_ttlmarker * this.keep_plug12_inch[kd].result_plug12_inch

        this.sum_plug12_inch = Number(this.sum_plug12_inch) + Number(this.result_sum_plug12_inch)
        
      }
     
      this.sum_plug12_inch = this.sum_plug12_inch / this.result_sum_ttlmarker
      
      this.sum_plug12_inch_last = parseFloat(this.sum_plug12_inch).toFixed(2)
       worksheet.getCell("K"+d).value =  Number(this.sum_plug12_inch_last);
       worksheet.getCell("K"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+d).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      this.result_sum_plug12_inch_l=0
      this.sum_plug12_inch_l =0
       for(var ld = 0; ld<this.keep_ttlmarker.length; ld++){
         this.result_sum_plug12_inch_l =  this.keep_perplug12[ld].result_perplug12 * this.keep_ttlmarker[ld].result_ttlmarker 
       
        this.sum_plug12_inch_l = Number(this.sum_plug12_inch_l) + Number(this.result_sum_plug12_inch_l)
    
      }
     
      this.sum_plug12_inch_l = this.sum_plug12_inch_l/this.result_sum_ttlmarker 
      
      //worksheet.getCell("L"+d).value = parseFloat(this.sum_plug12_inch_l).toFixed(2)+"%";
      
      worksheet.getCell("L"+d).value = this.sum_plug12_inch_l/100;
      worksheet.getCell("L"+d).numFmt = '0.00%';
      worksheet.getCell("L"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+d).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

  
          this.result_avg_end_plug12 = 0
          this.sum_avg_end_plug12 = 0
         for(var md = 0; md<this.keep_ttlmarker.length; md++){
          
         this.result_avg_end_plug12 = this.keep_avg_loss_plug12[md].all_result_avg * this.keep_ttlmarker[md].result_ttlmarker 
        
   
        this.sum_avg_end_plug12 = Number(this.sum_avg_end_plug12) + Number(this.result_avg_end_plug12)
        
      } 
     this.sum_avg_end_plug12 = this.sum_avg_end_plug12/this.result_sum_ttlmarker

     this.sum_avg_end_plug12_last = parseFloat(this.sum_avg_end_plug12).toFixed(2)
      worksheet.getCell("M"+d).value = Number(this.sum_avg_end_plug12_last);
      worksheet.getCell("M"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+d).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
  this.all_sum_length_yd = 0
   
     for(var azc = 0; azc<this.keep_avg_length_yd.length; azc++){
       this.all_sum_length_yd = Number(this.all_sum_length_yd)+ Number(this.keep_avg_length_yd[azc].all_result_avg_length)
      
      }   
    
      this.all_sum_length_yd_last = parseFloat(this.all_sum_length_yd).toFixed(2)


      
      worksheet.getCell("N"+d).value = Number(this.all_sum_length_yd_last);
      worksheet.getCell("M"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+d).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" }, 
        right: { style: "double" },
      };

        this.sum_avg_per = 0
        this.sum_avg_per = this.all_sum_length_yd_last / this.last_result_sum_ttlmarker * 100

        this.all_value_before_diff.push({
      value:this.sum_avg_per,
      value_name:"End Loss"
    })
        
       
  
       worksheet.getCell("O"+d).value = this.sum_avg_per/100;
       worksheet.getCell("O"+d).numFmt = '0.00%'; 
       worksheet.getCell("O"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+d).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      this.result_splice = 0
      for(var re_splice = 0; re_splice<this.keep_splice_Loss.length; re_splice++ ){
        this.result_splice =  Number(this.result_splice) + Number(this.keep_splice_Loss[re_splice].result_splice)
        
      }
    // this.result_splice = this.result_splice/resp.data.data.length
     this.keep_splice_Loss.push({
          result_splice:this.result_splice
     })

     this.result_splice = parseFloat(this.result_splice).toFixed(2)
       worksheet.getCell("P"+d).value = Number(this.result_splice);
       worksheet.getCell("P"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+d).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
 
      this.sum_result_splice_per = this.result_splice/this.result_sum_ttlmarker*100

      this.all_value_before_diff.push({
      value:this.sum_result_splice_per,
      value_name:"Splice Loss"
    })


      
      
       worksheet.getCell("Q"+d).value = this.sum_result_splice_per/100;
       worksheet.getCell("Q"+d).numFmt = '0.00%'; 
       worksheet.getCell("Q"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+d).border = {
        top: { style: "double" }, 
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      
      
       this.result_sum_cut_loss = 0
      for(var re_cut_out = 0; re_cut_out<this.keep_cut_out_loss.length; re_cut_out++ ){
        this.result_sum_cut_loss =  Number(this.result_sum_cut_loss) + Number(this.keep_cut_out_loss[re_cut_out].result_cut_out_loss)
        
      }
      
      this.result_sum_cut_loss_last = parseFloat(this.result_sum_cut_loss).toFixed(2)
       worksheet.getCell("R"+d).value =  Number(this.result_sum_cut_loss_last);
       worksheet.getCell("R"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+d).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
     
      this.sum_cut_loss_out_per = this.result_sum_cut_loss/this.result_sum_ttlmarker*100

      this.all_value_before_diff.push({
      value:this.sum_cut_loss_out_per,
      value_name:"Cut Out Loss"
    })
     
       worksheet.getCell("S"+d).value =  this.sum_cut_loss_out_per/100;
       worksheet.getCell("S"+d).numFmt = '0.00%'; 
       worksheet.getCell("S"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+d).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
 this.result_sum_remnants = 0
      for(var re_rem = 0; re_rem<this.keep_remnants.length; re_rem++ ){
        this.result_sum_remnants =  Number(this.result_sum_remnants) + Number(this.keep_remnants[re_rem].result_remnants)
        
        
      }

      this.result_sum_remnants_last = parseFloat(this.result_sum_remnants).toFixed(2)
       worksheet.getCell("T"+d).value = Number(this.result_sum_remnants_last);
       worksheet.getCell("T"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+d).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      this.sum_remnants_per = this.result_sum_remnants/this.result_sum_ttlmarker*100

      this.all_value_before_diff.push({
      value:this.sum_remnants_per,
      value_name:"Remnants Loss"
    })

      
       worksheet.getCell("U"+d).value = this.sum_remnants_per/100;
       worksheet.getCell("U"+d).numFmt = '0.00%'; 
       worksheet.getCell("U"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+d).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      this.all_sum_total = 0;
      for(var abc = 0; abc<this.keep_total_loss.length; abc++){
        
        this.all_sum_total = Number(this.all_sum_total) + Number(this.keep_total_loss[abc].result_total_length)
      }
      
      this.all_sum_total_last = parseFloat(this.all_sum_total).toFixed(2)
       worksheet.getCell("V"+d).value = Number(this.all_sum_total_last);
       worksheet.getCell("V"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+d).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };
      this.sum_total_length_per = this.all_sum_total/this.result_sum_ttlmarker*100
       worksheet.getCell("W"+d).value = this.sum_total_length_per/100;
       worksheet.getCell("W"+d).numFmt = '0.00%'; 
       worksheet.getCell("W"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+d).border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };






      

      this.minus_of_diff = 0
      for(var ax = 0; ax<this.all_value_before_diff.length; ax++){
        if(this.all_value_before_diff[ax].value == 0){
        this.result_minus_diff.push({
          value:"",
          value_main:"",
          value_name:this.all_value_before_target[ax].value_name,
          value_target:this.all_value_before_target[ax].value_target
        })
        }else{
        this.minus_of_diff = 0
        this.minus_of_diff = Number(this.all_value_before_target[ax].value) - Number(this.all_value_before_diff[ax].value)
        if(this.minus_of_diff < 0){
          this.minus_of_diff = 0-(this.minus_of_diff)
        }
        this.result_minus_diff.push({
          value:this.minus_of_diff,
          value_main:this.all_value_before_diff[ax].value,
          value_name:this.all_value_before_target[ax].value_name,
          value_target:this.all_value_before_target[ax].value_target
        })
        }

      
      }

       
   
        this.find_max= 0
        this.find_max_main = 0
        this.find_max_name =""
        this.find_max_target = ""
        this.count1 = 0
        for(ax = 0; ax<this.result_minus_diff.length; ax++){
          if(this.count1 = 0){
          if(this.result_minus_diff[ax].value > 0){
            this.find_max = this.result_minus_diff[ax].value
            this.find_max_main = this.result_minus_diff[ax].value_main
            this.find_max_name = this.result_minus_diff[ax].value_name
            this.find_max_target = this.result_minus_diff[ax].value_target
          }
          this.count1 ++
         }else if(this.result_minus_diff[ax].value == ""){
           this.find_max = this.find_max
            this.find_max_main = this.find_max_main
             this.find_max_name = this.find_max_name
             this.find_max_target = this.find_max_target
         }else{
           if(this.find_max < this.result_minus_diff[ax].value){
            this.find_max = this.result_minus_diff[ax].value
            this.find_max_main = this.result_minus_diff[ax].value_main
             this.find_max_name = this.result_minus_diff[ax].value_name
             this.find_max_target = this.result_minus_diff[ax].value_target
             
           }else if(this.find_max == 0){
             this.find_max = this.result_minus_diff[ax].value
             this.find_max_main = this.result_minus_diff[ax].value_main
             this.find_max_name = this.result_minus_diff[ax].value_name
             this.find_max_target = this.result_minus_diff[ax].value_target
           }
         }
         this.count1 ++
        }
 
     this.count_row = 0

  
      if(this.sum_total_length_per > this.result_overall_total){
        worksheet.getCell("L"+e).value = "หมายเหตุ  :  % Spread Loss โดยรวม = "+parseFloat(this.sum_total_length_per).toFixed(4)+ "% สูงกว่า Target  ("+parseFloat(this.result_overall_total).toFixed(4)+"%) ความสูญเสียเกิดจาก";
        
      }else if(this.sum_total_length_per == this.result_overall_total){
        worksheet.getCell("L"+e).value = "หมายเหตุ  :  % Spread Loss โดยรวม = "+parseFloat(this.sum_total_length_per).toFixed(2)+ "% เท่ากัน Target  ("+parseFloat(this.result_overall_total).toFixed(2)+"%) ความสูญเสียเกิดจาก";
     
      }else{
        worksheet.getCell("L"+e).value = "หมายเหตุ  :  % Spread Loss โดยรวม = "+parseFloat(this.sum_total_length_per).toFixed(4)+ "% ต่ำกว่า Target  ("+parseFloat(this.result_overall_total).toFixed(4)+"%) ความสูญเสียเกิดจาก";

      }

      if(this.all_value_before_diff[0].value > this.all_value_before_target[0].value){
       worksheet.getCell("N"+f).value = "%End Loss ="+ parseFloat(this.all_value_before_diff[0].value).toFixed(2)+"% (Target "+parseFloat(this.all_value_before_target[0].value).toFixed(2)+"%) ซึ่งเป็นจุดที่ควรควบคุม";
    
 
      }else{
          worksheet.getCell("N"+f).value = "%End Loss ="+ parseFloat(this.all_value_before_diff[0].value).toFixed(2)+"% (Target "+parseFloat(this.all_value_before_target[0].value).toFixed(2)+"%)";
    
   
      }

      if(this.all_value_before_diff[1].value > 0){
       
      if(this.all_value_before_diff[1].value > this.all_value_before_target[1].value){
       worksheet.getCell("N"+g).value = "%Splice Loss ="+ parseFloat(this.all_value_before_diff[1].value).toFixed(2)+"% (Target "+parseFloat(this.all_value_before_target[1].value).toFixed(2)+"%)ซึ่งเป็นจุดที่ควรควบคุม";
    
 
      }else{
         worksheet.getCell("N"+g).value = "%Splice Loss ="+ parseFloat(this.all_value_before_diff[1].value).toFixed(2)+"% (Target "+parseFloat(this.all_value_before_target[1].value).toFixed(2)+"%)";
    

      }
      }else{
        this.count_row++
      }

      if(this.all_value_before_diff[2].value > 0){
        
      if(this.all_value_before_diff[2].value > this.all_value_before_target[2].value){
       worksheet.getCell("N"+h).value = "%Cut-Out Loss ="+ parseFloat(this.all_value_before_diff[2].value).toFixed(2)+"% (Target "+parseFloat(this.all_value_before_target[2].value).toFixed(2)+"%)  ซึ่งเป็นจุดที่ควรควบคุม";
    

      }else{
         worksheet.getCell("N"+h).value = "%Cut-Out Loss ="+ parseFloat(this.all_value_before_diff[2].value).toFixed(2)+"% (Target "+parseFloat(this.all_value_before_target[2].value).toFixed(2)+"%)";

      }
      }else{
        this.count_row++
      }

      if(this.all_value_before_diff[3].value > 0){
    
      if(this.all_value_before_diff[3].value > this.all_value_before_target[3].value){
       worksheet.getCell("N"+i).value = "%Remnants Loss ="+ parseFloat(this.all_value_before_diff[3].value).toFixed(2)+"% (Target "+parseFloat(this.all_value_before_target[3].value).toFixed(2)+"%)  ซึ่งเป็นจุดที่ควรควบคุม";
    
  
      }else{
      worksheet.getCell("N"+i).value = "%Remnants Loss ="+ parseFloat(this.all_value_before_diff[3].value).toFixed(2)+"% (Target "+parseFloat(this.all_value_before_target[3].value).toFixed(2)+"%)";
    
  
      }
      }else{
        this.count_row++
      }
      
      j = Number(j) - Number(this.count_row)
      if(this.sum_plug12_inch_l > 1.50){
       worksheet.getCell("N"+j).value = "% Width Loss  โดยรวม ="+parseFloat(this.sum_plug12_inch_l).toFixed(2) +"%  สูงกว่า Target (1.50%) ซึ่งเป็นจุดที่ควรควบคุม";
    
      }else if(this.sum_plug12_inch_l == 1.50){
        worksheet.getCell("N"+j).value = "% Width Loss  โดยรวม ="+parseFloat(this.sum_plug12_inch_l).toFixed(2) +"%  เท่ากัน Target (1.50%)";
  
      }else{
         worksheet.getCell("N"+j).value = "% Width Loss  โดยรวม ="+parseFloat(this.sum_plug12_inch_l).toFixed(2) +"%  ต่ำกว่า Target (1.50%)";
  
      }

      
       worksheet.getCell("L"+k).value = "Target";
      worksheet.mergeCells("L"+k+":" +"O"+k);
      worksheet.getCell("L"+k).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+k).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       
       worksheet.getCell("P"+l).value = "Weighted target of this Week";
      worksheet.mergeCells("P"+l+":" +"P"+m);
      worksheet.getCell("P"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 

        worksheet.getCell("Q"+k).value = "Actual Performance";
      worksheet.mergeCells("Q"+k+":" +"V"+k);
      worksheet.getCell("Q"+k).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+k).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("L"+l).value = "Woven";
      worksheet.mergeCells("L"+l+":" +"L"+m);
      worksheet.getCell("L"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("Q"+l).value = "Woven";
      worksheet.mergeCells("Q"+l+":" +"Q"+m);
      worksheet.getCell("Q"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("M"+l).value = "Knit";
      worksheet.mergeCells("M"+l+":" +"O"+l);
      worksheet.getCell("M"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
        worksheet.getCell("R"+l).value = "Knit";
      worksheet.mergeCells("R"+l+":" +"T"+l);
      worksheet.getCell("R"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 


        worksheet.getCell("U"+l).value = "of This Week";
      worksheet.mergeCells("U"+l+":" +"V"+l);
      worksheet.getCell("U"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("M"+m).value = "Light & Mid Wt.";
       worksheet.getCell("M"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("N"+m).value = "Fleece";
       worksheet.getCell("N"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("O"+m).value = "Special";
       worksheet.getCell("O"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("P"+m).value = "Overall";
       worksheet.getCell("P"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
       worksheet.getCell("R"+m).value = "Light & Mid Wt.";
       worksheet.getCell("R"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("S"+m).value = "Fleece";
       worksheet.getCell("S"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("T"+m).value = "Special";
       worksheet.getCell("T"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("U"+m).value = "Average";
       worksheet.getCell("U"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("V"+m).value = "Total";
       worksheet.getCell("V"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
      

        worksheet.getCell("L"+n).value = "";
        worksheet.getCell("L"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
        worksheet.getCell("M"+n).value = "";
        worksheet.getCell("M"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


      worksheet.getCell("N"+n).value = "";
      worksheet.getCell("N"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("O"+n).value = "";
      worksheet.getCell("O"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("P"+n).value = "";
       worksheet.getCell("P"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
 

    
     worksheet.getCell("Q"+n).value = Number(this.total_overall[0].result_ttl);
      worksheet.getCell("Q"+n).numFmt = "#,##0.00";
      worksheet.getCell("Q"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
         };
     
        worksheet.getCell("R"+n).value = Number(this.total_overall[1].result_ttl);
        worksheet.getCell("R"+n).numFmt = "#,##0.00";
       worksheet.getCell("R"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("S"+n).value = Number(this.total_overall[2].result_ttl);
      worksheet.getCell("S"+n).numFmt = "#,##0.00";
        worksheet.getCell("S"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

           worksheet.getCell("T"+n).value = Number(this.total_overall[3].result_ttl);
           worksheet.getCell("T"+n).numFmt = "#,##0.00";
          worksheet.getCell("T"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
   
        
    
      
      this.devide_ttlmarker = this.result_sum_ttlmarker/this.keep_ttlmarker.length
         worksheet.getCell("U"+n).value = "";
         worksheet.getCell("U"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    
      worksheet.getCell("V"+n).value = Number(this.result_sum_ttlmarker);
      worksheet.getCell("V"+n).numFmt = "#,##0.00";
      worksheet.getCell("V"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
         worksheet.getCell("J"+o).value = "End Loss";
      worksheet.mergeCells("J"+o+":" +"K"+o);
      worksheet.getCell("J"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
        worksheet.getCell("L"+o).value = 0.23/100;
        worksheet.getCell("L"+o).numFmt = "0.00%";
      worksheet.getCell("L"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("M"+o).value = 0.32/100;
        worksheet.getCell("M"+o).numFmt = "0.00%";
      worksheet.getCell("M"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }, 
      };

      worksheet.getCell("N"+o).value = 0.35/100;
      worksheet.getCell("N"+o).numFmt = "0.00%";
      worksheet.getCell("N"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("O"+o).value = 0.45/100;
      worksheet.getCell("O"+o).numFmt = "0.00%";
      worksheet.getCell("O"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


      
      for(var xi = 0; xi<this.multiple.length; xi++){
       
      if(this.multiple[xi] == "Woven"){
        if(this.keep_ttlmarker.length == 4){
        this.overall_end_loss = 0.23*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_end_loss.push({
          total_endloss:this.overall_end_loss
        })
        }else if(this.keep_ttlmarker.length == 3){
          this.overall_end_loss = 0.23*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_end_loss.push({
          total_endloss:this.overall_end_loss
        })
        }else if(this.keep_ttlmarker.length == 2){
          this.overall_end_loss = 0.23*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_end_loss.push({
          total_endloss:this.overall_end_loss
        })
        }else if(this.keep_ttlmarker.length == 1){
          this.overall_end_loss = 0.23*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_end_loss.push({
          total_endloss:this.overall_end_loss
        })
        }
  
      } else if(this.multiple[xi]=="Light-Middle Weight"){
        if(this.keep_ttlmarker.length == 4){
          this.overall_end_loss = 0.32*this.keep_ttlmarker[1].result_ttlmarker
          this.over_all_end_loss.push({
          total_endloss:this.overall_end_loss
        })
      }else if(this.keep_ttlmarker.length == 3){
        this.overall_end_loss = 0.32*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_end_loss.push({
          total_endloss:this.overall_end_loss
        })
      
      }else if(this.keep_ttlmarker.length == 2){
        this.overall_end_loss = 0.32*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_end_loss.push({
          total_endloss:this.overall_end_loss
        })

      }else if(this.keep_ttlmarker.length == 1){
        this.overall_end_loss = 0.32*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_end_loss.push({
          total_endloss:this.overall_end_loss
        })
      }

      }else if(this.multiple[xi] == "Fleece"){
           if(this.keep_ttlmarker.length == 4){
          this.overall_end_loss = 0.35*this.keep_ttlmarker[2].result_ttlmarker
          this.over_all_end_loss.push({
          total_endloss:this.overall_end_loss
        })
      }else if(this.keep_ttlmarker.length == 3){
        this.overall_end_loss = 0.35*this.keep_ttlmarker[1].result_ttlmarker
        this.over_all_end_loss.push({
          total_endloss:this.overall_end_loss
        })
      
      }else if(this.keep_ttlmarker.length == 2){
        this.overall_end_loss = 0.35*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_end_loss.push({
          total_endloss:this.overall_end_loss
        })

      }else if(this.keep_ttlmarker.length == 1){
        this.overall_end_loss = 0.32*this.keep_ttlmarker[0].result_ttlmarker
        
        this.over_all_end_loss.push({
          total_endloss:this.overall_end_loss
        })
      }
      }else if(this.multiple[xi] == "Special"){
           if(this.keep_ttlmarker.length == 4){
          this.overall_end_loss = 0.45*this.keep_ttlmarker[3].result_ttlmarker
          this.over_all_end_loss.push({
          total_endloss:this.overall_end_loss
        })
      }else if(this.keep_ttlmarker.length == 3){
        this.overall_end_loss = 0.45*this.keep_ttlmarker[2].result_ttlmarker
        this.over_all_end_loss.push({
          total_endloss:this.overall_end_loss
        })
      
      }else if(this.keep_ttlmarker.length == 2){
        this.overall_end_loss = 0.45*this.keep_ttlmarker[1].result_ttlmarker
        this.over_all_end_loss.push({
          total_endloss:this.overall_end_loss
        })

      }else if(this.keep_ttlmarker.length == 1){
        this.overall_end_loss = 0.45*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_end_loss.push({
          total_endloss:this.overall_end_loss
        })
      }
      }
      }
     
      //L42
      //Q41 = this.keep_ttlmarker[0].result_ttlmarker
      //all = this.result_sum_ttlmarker
  
    
    
    
   
   
      worksheet.getCell("P"+o).value = this.result_overall_end_loss/100;
      worksheet.getCell("P"+o).numFmt = '0.00%'; 
      worksheet.getCell("P"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };



        worksheet.getCell("Q"+o).value = this.end_loss_type[0].result_endloss/100
        worksheet.getCell("Q"+o).numFmt = '0.00%'; 
      worksheet.getCell("Q"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

             worksheet.getCell("R"+o).value = this.end_loss_type[1].result_endloss/100
             worksheet.getCell("R"+o).numFmt = '0.00%'; 
      worksheet.getCell("R"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
        worksheet.getCell("S"+o).value =  this.end_loss_type[2].result_endloss/100
        worksheet.getCell("S"+o).numFmt = '0.00%'; 
      worksheet.getCell("S"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

         worksheet.getCell("T"+o).value =  this.end_loss_type[3].result_endloss/100
         worksheet.getCell("T"+o).numFmt = '0.00%'; 
      worksheet.getCell("T"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
   
      worksheet.getCell("U"+o).value = "";
      worksheet.getCell("U"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("V"+o).value = "";
      worksheet.getCell("V"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      


        worksheet.getCell("J"+p).value = "Splice Loss";
      worksheet.mergeCells("J"+p+":" +"K"+p);
      worksheet.getCell("J"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
        worksheet.getCell("L"+p).value = 0.10/100;
        worksheet.getCell("L"+p).numFmt = '0.00%'; 
      worksheet.getCell("L"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("M"+p).value = 0.10/100;
        worksheet.getCell("M"+p).numFmt = '0.00%'; 
      worksheet.getCell("M"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("N"+p).value = 0.10/100;
      worksheet.getCell("N"+p).numFmt = '0.00%'; 
      worksheet.getCell("N"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("O"+p).value = 0.10/100;
      worksheet.getCell("O"+p).numFmt = '0.00%'; 
      worksheet.getCell("O"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.overall_splice_loss = 0
      for(var xi = 0; xi<this.multiple.length; xi++){
       
      if(this.multiple[xi] == "Woven"){
        if(this.keep_ttlmarker.length == 4){
        this.overall_splice_loss = 0.10*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_splice_loss.push({
          total_splice_loss:this.overall_splice_loss
        })
        }else if(this.keep_ttlmarker.length == 3){
          this.overall_splice_loss = 0.10*this.keep_ttlmarker[0].result_ttlmarker
       this.over_all_splice_loss.push({
          total_splice_loss:this.overall_splice_loss
        })
        }else if(this.keep_ttlmarker.length == 2){
          this.overall_splice_loss = 0.10*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_splice_loss.push({
          total_splice_loss:this.overall_splice_loss
        })
        }else if(this.keep_ttlmarker.length == 1){
          this.overall_splice_loss = 0.10*this.keep_ttlmarker[0].result_ttlmarker
       this.over_all_splice_loss.push({
          total_splice_loss:this.overall_splice_loss
        })
        }
  
      } else if(this.multiple[xi]=="Light-Middle Weight"){
        if(this.keep_ttlmarker.length == 4){
          this.overall_splice_loss = 0.10*this.keep_ttlmarker[1].result_ttlmarker
         this.over_all_splice_loss.push({
          total_splice_loss:this.overall_splice_loss
        })
      }else if(this.keep_ttlmarker.length == 3){
        this.overall_splice_loss = 0.10*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_splice_loss.push({
          total_splice_loss:this.overall_splice_loss
        })
      
      }else if(this.keep_ttlmarker.length == 2){
        this.overall_splice_loss = 0.10*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_splice_loss.push({
          total_splice_loss:this.overall_splice_loss
        })

      }else if(this.keep_ttlmarker.length == 1){
        this.overall_splice_loss = 0.10*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_splice_loss.push({
          total_splice_loss:this.overall_splice_loss
        })
      }

      }else if(this.multiple[xi] == "Fleece"){
           if(this.keep_ttlmarker.length == 4){
          this.overall_splice_loss = 0.10*this.keep_ttlmarker[2].result_ttlmarker
          this.over_all_splice_loss.push({
          total_splice_loss:this.overall_splice_loss
        })
      }else if(this.keep_ttlmarker.length == 3){
        this.overall_splice_loss = 0.10*this.keep_ttlmarker[1].result_ttlmarker
        this.over_all_splice_loss.push({
          total_splice_loss:this.overall_splice_loss
        })
      
      }else if(this.keep_ttlmarker.length == 2){
        this.overall_splice_loss = 0.10*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_splice_loss.push({
          total_splice_loss:this.overall_splice_loss
        })

      }else if(this.keep_ttlmarker.length == 1){
        this.overall_splice_loss = 0.10*this.keep_ttlmarker[0].result_ttlmarker
       this.over_all_splice_loss.push({
          total_splice_loss:this.overall_splice_loss
        })
      }
      }else if(this.multiple[xi] == "Special"){
           if(this.keep_ttlmarker.length == 4){
          this.overall_splice_loss = 0.10*this.keep_ttlmarker[3].result_ttlmarker
         this.over_all_splice_loss.push({
          total_splice_loss:this.overall_splice_loss
        })
      }else if(this.keep_ttlmarker.length == 3){
        this.overall_splice_loss = 0.10*this.keep_ttlmarker[2].result_ttlmarker
       this.over_all_splice_loss.push({
          total_splice_loss:this.overall_splice_loss
        })
      
      }else if(this.keep_ttlmarker.length == 2){
        this.overall_splice_loss = 0.10*this.keep_ttlmarker[1].result_ttlmarker
        this.over_all_splice_loss.push({
          total_splice_loss:this.overall_splice_loss
        })

      }else if(this.keep_ttlmarker.length == 1){
        this.overall_splice_loss = 0.10*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_splice_loss.push({
          total_splice_loss:this.overall_splice_loss
        })
      }
      }
      }
   
    
      worksheet.getCell("P"+p).value = this.result_overall_splice_loss/100
      worksheet.getCell("P"+p).numFmt = '0.00%'; 
      worksheet.getCell("P"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };



        worksheet.getCell("Q"+p).value = this.splice_loss_type[0].result_splice_loss/100
        worksheet.getCell("Q"+p).numFmt = '0.00%'; 
      worksheet.getCell("Q"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("R"+p).value = this.splice_loss_type[1].result_splice_loss/100
       worksheet.getCell("R"+p).numFmt = '0.00%'; 
      worksheet.getCell("R"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
  };

      worksheet.getCell("S"+p).value =  this.splice_loss_type[2].result_splice_loss/100
      worksheet.getCell("S"+p).numFmt = '0.00%'; 
      worksheet.getCell("S"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },

  };

      worksheet.getCell("T"+p).value =  this.splice_loss_type[3].result_splice_loss/100
      worksheet.getCell("T"+p).numFmt = '0.00%'; 
      worksheet.getCell("T"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


      worksheet.getCell("U"+p).value = "";
      worksheet.getCell("U"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("V"+p).value = "";
      worksheet.getCell("V"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      

      worksheet.getCell("J"+q).value = "Cut - Out Loss";
      worksheet.mergeCells("J"+q+":" +"K"+q);
      worksheet.getCell("J"+q).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+q).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("L"+q).value = 0.00/100;
       worksheet.getCell("L"+q).numFmt = "0.00%";
      worksheet.getCell("L"+q).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+q).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("M"+q).value = 0.00/100;
      worksheet.getCell("M"+q).numFmt = "0.00%";
      worksheet.getCell("M"+q).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+q).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("N"+q).value = 0.00/100;
      worksheet.getCell("N"+q).numFmt = "0.00%";
      worksheet.getCell("N"+q).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+q).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("O"+q).value = 0.00/100;
      worksheet.getCell("O"+q).numFmt = "0.00%";
      worksheet.getCell("O"+q).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+q).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.overall_cut_loss = 0
      for(var xi = 0; xi<this.multiple.length; xi++){
       
      if(this.multiple[xi] == "Woven"){
        if(this.keep_ttlmarker.length == 4){
        this.overall_cut_loss = 0.00*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_cut_out_loss.push({
          total_cut_loss:this.overall_cut_loss
        })
        }else if(this.keep_ttlmarker.length == 3){
          this.overall_cut_loss = 0.00*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_cut_out_loss.push({
          total_cut_loss:this.overall_cut_loss
        })
        }else if(this.keep_ttlmarker.length == 2){
          this.overall_cut_loss = 0.00*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_cut_out_loss.push({
          total_cut_loss:this.overall_cut_loss
        })
        }else if(this.keep_ttlmarker.length == 1){
          this.overall_cut_loss = 0.00*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_cut_out_loss.push({
          total_cut_loss:this.overall_cut_loss
        })
        }
  
      } else if(this.multiple[xi]=="Light-Middle Weight"){
        if(this.keep_ttlmarker.length == 4){
          this.overall_cut_loss = 0.00*this.keep_ttlmarker[1].result_ttlmarker
          this.over_all_cut_out_loss.push({
          total_cut_loss:this.overall_cut_loss
        })
      }else if(this.keep_ttlmarker.length == 3){
        this.overall_cut_loss = 0.00*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_cut_out_loss.push({
          total_cut_loss:this.overall_cut_loss
        })
      
      }else if(this.keep_ttlmarker.length == 2){
        this.overall_cut_loss = 0.00*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_cut_out_loss.push({
          total_cut_loss:this.overall_cut_loss
        })

      }else if(this.keep_ttlmarker.length == 1){
        this.overall_cut_loss = 0.00*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_cut_out_loss.push({
          total_cut_loss:this.overall_cut_loss
        })
      }

      }else if(this.multiple[xi] == "Fleece"){
           if(this.keep_ttlmarker.length == 4){
          this.overall_cut_loss = 0.00*this.keep_ttlmarker[2].result_ttlmarker
          this.over_all_cut_out_loss.push({
          total_cut_loss:this.overall_cut_loss
        })
      }else if(this.keep_ttlmarker.length == 3){
        this.overall_cut_loss = 0.00*this.keep_ttlmarker[1].result_ttlmarker
        this.over_all_cut_out_loss.push({
          total_cut_loss:this.overall_cut_loss
        })
      
      }else if(this.keep_ttlmarker.length == 2){
        this.overall_cut_loss = 0.00*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_cut_out_loss.push({
          total_cut_loss:this.overall_cut_loss
        })

      }else if(this.keep_ttlmarker.length == 1){
        this.overall_cut_loss = 0.00*this.keep_ttlmarker[0].result_ttlmarker
        
        this.over_all_cut_out_loss.push({
          total_cut_loss:this.overall_cut_loss
        })
      }
      }else if(this.multiple[xi] == "Special"){
           if(this.keep_ttlmarker.length == 4){
          this.overall_cut_loss = 0.00*this.keep_ttlmarker[3].result_ttlmarker
          this.over_all_cut_out_loss.push({
          total_cut_loss:this.overall_cut_loss
        })
      }else if(this.keep_ttlmarker.length == 3){
        this.overall_cut_loss = 0.00*this.keep_ttlmarker[2].result_ttlmarker
        this.over_all_cut_out_loss.push({
          total_cut_loss:this.overall_cut_loss
        })
      
      }else if(this.keep_ttlmarker.length == 2){
        this.overall_cut_loss = 0.00*this.keep_ttlmarker[1].result_ttlmarker
        this.over_all_cut_out_loss.push({
          total_cut_loss:this.overall_cut_loss
        })

      }else if(this.keep_ttlmarker.length == 1){
        this.overall_cut_loss = 0.00*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_cut_out_loss.push({
          total_cut_loss:this.overall_cut_loss
        })
      }
      }
      }


   
      worksheet.getCell("P"+q).value = this.result_overall_cut/100
      worksheet.getCell("P"+q).numFmt = '0.00%'; 
      worksheet.getCell("P"+q).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+q).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("Q"+q).value = this.cut_out_loss_type[0].result_cut_out_loss/100
        worksheet.getCell("Q"+q).numFmt = '0.00%'; 
      worksheet.getCell("Q"+q).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+q).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

          worksheet.getCell("R"+q).value = this.cut_out_loss_type[1].result_cut_out_loss/100
          worksheet.getCell("R"+q).numFmt = '0.00%'; 
      worksheet.getCell("R"+q).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+q).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
 };
           worksheet.getCell("S"+q).value =  this.cut_out_loss_type[2].result_cut_out_loss/100
           worksheet.getCell("S"+q).numFmt = '0.00%'; 
      worksheet.getCell("S"+q).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+q).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
 };

        worksheet.getCell("T"+q).value = this.cut_out_loss_type[3].result_cut_out_loss/100
        worksheet.getCell("T"+q).numFmt = '0.00%'; 
      worksheet.getCell("T"+q).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+q).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    

      worksheet.getCell("U"+q).value = "";
      worksheet.getCell("U"+q).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+q).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("V"+q).value = "";
      worksheet.getCell("V"+q).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+q).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      

        worksheet.getCell("J"+r).value = "Remnants Loss";
      worksheet.mergeCells("J"+r+":" +"K"+r);
      worksheet.getCell("J"+r).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+r).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
        worksheet.getCell("L"+r).value = 0.05/100;
        worksheet.getCell("L"+r).numFmt = '0.00%'; 
      worksheet.getCell("L"+r).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+r).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("M"+r).value = 0.08/100;
        worksheet.getCell("M"+r).numFmt = '0.00%'; 
      worksheet.getCell("M"+r).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+r).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("N"+r).value = 0.11/100;
      worksheet.getCell("N"+r).numFmt = '0.00%'; 
      worksheet.getCell("N"+r).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+r).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("O"+r).value = 0.10/100;
      worksheet.getCell("O"+r).numFmt = '0.00%'; 
      worksheet.getCell("O"+r).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+r).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

         this.overall_remnants = 0
      for(var xi = 0; xi<this.multiple.length; xi++){
      
      if(this.multiple[xi] == "Woven"){
        if(this.keep_ttlmarker.length == 4){
        this.overall_remnants = 0.05*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_remnants.push({
          total_remnants:this.overall_remnants
        })
        }else if(this.keep_ttlmarker.length == 3){
          this.overall_remnants = 0.05*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_remnants.push({
          total_remnants:this.overall_remnants
        })
        }else if(this.keep_ttlmarker.length == 2){
          this.overall_remnants = 0.05*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_remnants.push({
          total_remnants:this.overall_remnants
        })
        }else if(this.keep_ttlmarker.length == 1){
          this.overall_remnants = 0.05*this.keep_ttlmarker[0].result_ttlmarker
       this.over_all_remnants.push({
          total_remnants:this.overall_remnants
        })
        }
  
      } else if(this.multiple[xi]=="Light-Middle Weight"){
        if(this.keep_ttlmarker.length == 4){
          this.overall_remnants = 0.08*this.keep_ttlmarker[1].result_ttlmarker
          this.over_all_remnants.push({
          total_remnants:this.overall_remnants
        })
      }else if(this.keep_ttlmarker.length == 3){
        this.overall_remnants = 0.08*this.keep_ttlmarker[0].result_ttlmarker
          this.over_all_remnants.push({
          total_remnants:this.overall_remnants
        })
      
      }else if(this.keep_ttlmarker.length == 2){
        this.overall_remnants = 0.08*this.keep_ttlmarker[0].result_ttlmarker
          this.over_all_remnants.push({
          total_remnants:this.overall_remnants
        })

      }else if(this.keep_ttlmarker.length == 1){
        this.overall_remnants = 0.08*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_remnants.push({
          total_remnants:this.overall_remnants
        })
      }

      }else if(this.multiple[xi] == "Fleece"){
           if(this.keep_ttlmarker.length == 4){
          this.overall_remnants = 0.11*this.keep_ttlmarker[2].result_ttlmarker
          this.over_all_remnants.push({
          total_remnants:this.overall_remnants
        })
      }else if(this.keep_ttlmarker.length == 3){
        this.overall_remnants = 0.11*this.keep_ttlmarker[1].result_ttlmarker
        this.over_all_remnants.push({
          total_remnants:this.overall_remnants
        })
      
      }else if(this.keep_ttlmarker.length == 2){
        this.overall_remnants = 0.11*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_remnants.push({
          total_remnants:this.overall_remnants
        })

      }else if(this.keep_ttlmarker.length == 1){
       this.overall_remnants = 0.11*this.keep_ttlmarker[0].result_ttlmarker
       this.over_all_remnants.push({
          total_remnants:this.overall_remnants
        })
      }
      }else if(this.multiple[xi] == "Special"){
           if(this.keep_ttlmarker.length == 4){
          this.overall_remnants = 0.10*this.keep_ttlmarker[3].result_ttlmarker
          this.over_all_remnants.push({
          total_remnants:this.overall_remnants
        })
      }else if(this.keep_ttlmarker.length == 3){
        this.overall_remnants = 0.10*this.keep_ttlmarker[2].result_ttlmarker
       this.over_all_remnants.push({
          total_remnants:this.overall_remnants
        })
      
      }else if(this.keep_ttlmarker.length == 2){
        this.overall_remnants = 0.10*this.keep_ttlmarker[1].result_ttlmarker
        this.over_all_remnants.push({
          total_remnants:this.overall_remnants
        })

      }else if(this.keep_ttlmarker.length == 1){
        this.overall_remnants = 0.10*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_remnants.push({
          total_remnants:this.overall_remnants
        })
      }
      }
      }

      
    
    
      worksheet.getCell("P"+r).value = this.result_overall_rem/100
      worksheet.getCell("P"+r).numFmt = '0.00%'; 
      worksheet.getCell("P"+r).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+r).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

            worksheet.getCell("Q"+r).value = this.rem_loss_type[0].result_rem_loss/100
      worksheet.getCell("Q"+r).numFmt = '0.00%'; 
      worksheet.getCell("Q"+r).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+r).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },

         };

                worksheet.getCell("R"+r).value = this.rem_loss_type[1].result_rem_loss/100
                worksheet.getCell("R"+r).numFmt = '0.00%'; 
      worksheet.getCell("R"+r).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+r).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
  };

       worksheet.getCell("S"+r).value =  this.rem_loss_type[2].result_rem_loss/100
       worksheet.getCell("S"+r).numFmt = '0.00%'; 
      worksheet.getCell("S"+r).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+r).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

                 worksheet.getCell("T"+r).value = this.rem_loss_type[3].result_rem_loss/100
                 worksheet.getCell("T"+r).numFmt = '0.00%'; 
      worksheet.getCell("T"+r).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+r).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
      
      worksheet.getCell("U"+r).value = "";
      worksheet.getCell("U"+r).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+r).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("V"+r).value = "";
      worksheet.getCell("V"+r).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+r).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
     
        worksheet.getCell("J"+s).value = "Total Loss";
      worksheet.mergeCells("J"+s+":" +"K"+s);
      worksheet.getCell("J"+s).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+s).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    
        worksheet.getCell("L"+s).value = this.total_length_value[0].p1/100;
        worksheet.getCell("L"+s).numFmt = '0.00%'; 
      worksheet.getCell("L"+s).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+s).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("M"+s).value = this.total_length_value[0].p2/100;
        worksheet.getCell("M"+s).numFmt = '0.00%'; 
      worksheet.getCell("M"+s).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+s).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("N"+s).value = this.total_length_value[0].p3/100;
      worksheet.getCell("N"+s).numFmt = '0.00%'; 
      worksheet.getCell("N"+s).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+s).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("O"+s).value = this.total_length_value[0].p4/100
      worksheet.getCell("O"+s).numFmt = '0.00%'; 
      worksheet.getCell("O"+s).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+s).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    

      
         this.overall_total_loss = 0
      for(var xi = 0; xi<this.multiple.length; xi++){
      
      if(this.multiple[xi] == "Woven"){
        if(this.keep_ttlmarker.length == 4){
        this.overall_total_loss = 0.38*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_total_loss.push({
          total_loss:this.overall_total_loss
        })
        }else if(this.keep_ttlmarker.length == 3){
          this.overall_total_loss = 0.38*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_total_loss.push({
          total_loss:this.overall_total_loss
        })
        }else if(this.keep_ttlmarker.length == 2){
          this.overall_total_loss = 0.38*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_total_loss.push({
          total_loss:this.overall_total_loss
        })
        }else if(this.keep_ttlmarker.length == 1){
          this.overall_total_loss = 0.38*this.keep_ttlmarker[0].result_ttlmarker
       this.over_all_total_loss.push({
          total_loss:this.overall_total_loss
        })
        }
  
      } else if(this.multiple[xi]=="Light-Middle Weight"){
        if(this.keep_ttlmarker.length == 4){
          this.overall_total_loss = 0.50*this.keep_ttlmarker[1].result_ttlmarker
         this.over_all_total_loss.push({
          total_loss:this.overall_total_loss
        })
      }else if(this.keep_ttlmarker.length == 3){
        this.overall_total_loss = 0.50*this.keep_ttlmarker[0].result_ttlmarker
         this.over_all_total_loss.push({
          total_loss:this.overall_total_loss
        })
      
      }else if(this.keep_ttlmarker.length == 2){
        this.overall_total_loss = 0.50*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_total_loss.push({
          total_loss:this.overall_total_loss
        })

      }else if(this.keep_ttlmarker.length == 1){
        this.overall_total_loss = 0.50*this.keep_ttlmarker[0].result_ttlmarker
       this.over_all_total_loss.push({
          total_loss:this.overall_total_loss
        })
      }

      }else if(this.multiple[xi] == "Fleece"){
           if(this.keep_ttlmarker.length == 4){
          this.overall_total_loss = 0.56*this.keep_ttlmarker[2].result_ttlmarker
         this.over_all_total_loss.push({
          total_loss:this.overall_total_loss
        })
      }else if(this.keep_ttlmarker.length == 3){
        this.overall_total_loss = 0.56*this.keep_ttlmarker[1].result_ttlmarker
        this.over_all_total_loss.push({
          total_loss:this.overall_total_loss
        })
      
      }else if(this.keep_ttlmarker.length == 2){
        this.overall_total_loss = 0.56*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_total_loss.push({
          total_loss:this.overall_total_loss
        })

      }else if(this.keep_ttlmarker.length == 1){
       this.overall_total_loss = 0.56*this.keep_ttlmarker[0].result_ttlmarker
       this.over_all_total_loss.push({
          total_loss:this.overall_total_loss
        })
      }
      }else if(this.multiple[xi] == "Special"){
           if(this.keep_ttlmarker.length == 4){
          this.overall_total_loss = 0.65*this.keep_ttlmarker[3].result_ttlmarker
       this.over_all_total_loss.push({
          total_loss:this.overall_total_loss
        })
      }else if(this.keep_ttlmarker.length == 3){
        this.overall_total_loss = 0.65*this.keep_ttlmarker[2].result_ttlmarker
       this.over_all_total_loss.push({
          total_loss:this.overall_total_loss
        })
      
      }else if(this.keep_ttlmarker.length == 2){
        this.overall_total_loss = 0.65*this.keep_ttlmarker[1].result_ttlmarker
       this.over_all_total_loss.push({
          total_loss:this.overall_total_loss
        })

      }else if(this.keep_ttlmarker.length == 1){
        this.overall_total_loss = 0.65*this.keep_ttlmarker[0].result_ttlmarker
        this.over_all_total_loss.push({
          total_loss:this.overall_total_loss
        })
      }
      }
      }

  
     
      worksheet.getCell("P"+s).value = this.result_overall_total/100
      worksheet.getCell("P"+s).numFmt = '0.00%'; 
      worksheet.getCell("P"+s).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+s).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.sum_all_woven = 0
      
          this.sum_all_woven = Number(this.end_loss_type[0].result_endloss) + Number(this.splice_loss_type[0].result_splice_loss)
          +Number(this.cut_out_loss_type[0].result_cut_out_loss) + Number(this.rem_loss_type[0].result_rem_loss)
      
      worksheet.getCell("Q"+s).value = this.sum_all_woven/100
      worksheet.getCell("Q"+s).numFmt = '0.00%'; 
      worksheet.getCell("Q"+s).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+s).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.sum_all_light = 0
      this.sum_all_light = Number(this.end_loss_type[1].result_endloss) + Number(this.splice_loss_type[1].result_splice_loss)
          +Number(this.cut_out_loss_type[1].result_cut_out_loss) + Number(this.rem_loss_type[1].result_rem_loss)

                worksheet.getCell("R"+s).value = this.sum_all_light/100
                worksheet.getCell("R"+s).numFmt = '0.00%'; 
      worksheet.getCell("R"+s).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+s).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.sum_all_fleece = 0
      this.sum_all_fleece = Number(this.end_loss_type[2].result_endloss) + Number(this.splice_loss_type[2].result_splice_loss)
          +Number(this.cut_out_loss_type[2].result_cut_out_loss) + Number(this.rem_loss_type[2].result_rem_loss)

      worksheet.getCell("S"+s).value =  this.sum_all_fleece/100
      worksheet.getCell("S"+s).numFmt = '0.00%'; 
      worksheet.getCell("S"+s).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+s).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.sum_all_special = 0
      this.sum_all_special = Number(this.end_loss_type[3].result_endloss) + Number(this.splice_loss_type[3].result_splice_loss)
          +Number(this.cut_out_loss_type[3].result_cut_out_loss) + Number(this.rem_loss_type[3].result_rem_loss)


     worksheet.getCell("T"+s).value =  this.sum_all_special/100
     worksheet.getCell("T"+s).numFmt = '0.00%'; 
      worksheet.getCell("T"+s).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+s).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };



     

      worksheet.getCell("U"+s).value = "";
      worksheet.getCell("U"+s).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+s).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("V"+s).value = "";
      worksheet.getCell("V"+s).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+s).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
      this.multiple_chose = "ALL"

 
       workbook.xlsx.writeBuffer().then((data) => {
        const blob = new Blob([data], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8",
        });
        saveAs(blob, "Spread Loss Daily Repot NY"+org+" ("+this.start_date+").xlsx");
      });  
       
         }
         
       
      })
       
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

</style>