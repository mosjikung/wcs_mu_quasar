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
       multiple: [],
       multiple_chose:ref("ALL"),
       org:'',
       options:['ALL','Woven','Light-Middle Weight','Fleece','Special'],
        rows:6,
        start:'',
        end:'',
        start_date:'',
        end_date:'',
        data_count:10,
        rowexport:[],
        x:0,
        end_loss_plug12:[],
        keep_ttlmarker:[],
        keep_avg_end:[],
        keep_splice_Loss:[],
        keep_cut_out_loss:[],
        keep_remnants:[],
        keep_total_loss:[],
        keep_plug12_inch:[],
        keep_sum_x:[],
        keep_sum_ttl_x:[],
        keep_line_l:[],
        keep_avg_loss_plug12:[],
        keep_avg_marker_length:[],
        keep_avg_length_yd:[],
        keep_sum_length_splice_loss:[],
        keep_sum_cut_loss_out:[],
        keep_perplug12:[],
        keep_total_loss_per:[],
        end_loss_value:[],
        splice_value:[],
        cut_out_value:[],
        remnants_loss_value:[],




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
      
     
      
  },
    methods:{
      checkdata(){
        
        
          for(var x = 0; x<this.multiple.length; x++){
            
          }
      },
      async  exportexcel(){

        
        this.multiple=[]
        this.rowexport=[]
        this.end_loss_plug12=[]
        this.keep_ttlmarker=[]
        this.keep_avg_end=[]
        this.keep_splice_Loss=[]
        this.keep_cut_out_loss=[]
        this.keep_remnants=[]
        this.keep_total_loss=[]
        this.keep_plug12_inch=[]
        this.keep_line_l=[]
        this.keep_avg_loss_plug12=[]
        this.keep_avg_length_yd=[]
        this.keep_sum_length_splice_loss=[]
        this.keep_sum_cut_loss_out=[]
        this.keep_perplug12=[]
        this.keep_splice_loss_per=[]
        this.keep_avg_marker_length=[]
        this.keep_cut_out_loss_per=[]
        this.keep_remnants_per=[]
        this.keep_total_loss_per=[]
        this.check_value=''
        this.keep_actual=[]
        this.export_out=[]
        this.keep_sum_x=[]
        this.keep_sum_ttl_x=[]
           
      var x = 0
          const workbook = new Excel.Workbook();
      workbook.creator = "Nanyang";
      workbook.lastModifiedBy = "Admin";
      workbook.created = new Date(2021, 8, 30);
      workbook.modified = new Date();
      workbook.lastPrinted = new Date(2021, 7, 27);
      const worksheet = workbook.addWorksheet("New Sheet");

       worksheet.columns = [ {key: 'A', width: 10},{key: 'B', width: 7.63},{key: 'C', width: 40},{key: 'D', width: 30},{key: 'E', width: 70}
       ,{key: 'F', width: 30},{key: 'G', width: 8.7},{key: 'H', width: 6.7},{key: 'I', width: 6.7},{key: 'J', width: 12},{key: 'K', width: 12},
       {key: 'L', width: 6.7},{key: 'M', width: 6.7},{key: 'N', width: 6.7},{key: 'O', width: 6.7},{key: 'P', width: 6.7},{key: 'Q', width: 6.7},
        {key: 'R', width: 6.7},{key: 'S', width: 12},{key: 'T', width: 6.7},{key: 'U', width: 6.7},{key: 'V', width: 6.7},{key: 'W', width: 6.7},
        {key: 'X', width: 6.7},{key: 'AC', width: 15.13 }];


      let org = this.$q.localStorage.getItem("org");
       worksheet.getCell("A1").value = "SPREAD LOSS AUDIT WEEKLY DETAIL REPORT ( MU-SLA-2 ) : NY"+org //ส่วนหัว
     
      worksheet.getCell('A1').font = {  
        name: 'Angsana New',
        color: { argb: 'FF0120E5' },
        family: 4,
        size: 20,
        bold: true
      };
    
     worksheet.getCell("A2").value = "Period :"; //ส่วนหัว
     worksheet.getCell("A2").alignment = { horizontal: "center",vertical:"middle" };
   
      worksheet.getCell("B2").value = this.start_date; //ส่วนหัว
   
      worksheet.getCell("C2").value = "to"; //ส่วนหัว
      worksheet.getCell("C2").alignment = { horizontal: "center",vertical:"middle" };
  

      worksheet.getCell("D2").value = this.end_date; //ส่วนหัว
   
     worksheet.getCell("A3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("B3").value = "";
      worksheet.mergeCells("B3:D3");
      worksheet.getCell("B3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("B3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("E3").value = "Fabric Type";
      worksheet.mergeCells("E3:G3");
      worksheet.getCell("E3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("E3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("H3").value = "Marker";
      worksheet.mergeCells("H3:K3");
      worksheet.getCell("H3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("L3").value = "Width Loss";
      worksheet.mergeCells("L3:M3");
      worksheet.getCell("L3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("N3").value = "AVG. End Loss";
      worksheet.mergeCells("N3:O3");
      worksheet.getCell("N3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("P3").value = "Splice Loss";
      worksheet.mergeCells("P3:Q3");
      worksheet.getCell("P3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("R3").value = "Cut Loss Out";
      worksheet.mergeCells("R3:S3");
      worksheet.getCell("R3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("T3").value = "Remnants Loss";
      worksheet.mergeCells("T3:U3");
      worksheet.getCell("T3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("V3").value = "Total Length Loss";
      worksheet.mergeCells("V3:W3");
      worksheet.getCell("V3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
          var a = 4 // first row head column Date per table
          var b = 5 //mergeA
           var count = 0


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
        params20.append("end",this.end_date)
        params20.append("org", org);
        for (var pair of params20.entries()) {
                  console.log(pair[0] + ", " + pair[1]);
                }
       await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/check_data_have_weekly',
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

       for( x=0; x<this.multiple.length;  x++){
         
         
          const params = new FormData()

        params.append("start",this.start_date)
        params.append("end",this.end_date)
        
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
        url: this.$api_url + '/mainso.php/export_data_3',
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
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
      var ty = this.length
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
       worksheet.getCell("A"+ty).value = "Total Type :" + this.multiple_rec;
      worksheet.mergeCells("A"+ty +":" + "I"+ty);
      worksheet.getCell("A"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
        this.result_ttlmarker = 0
      for(var re_ttlmarker = 0; re_ttlmarker<resp.data.data.length; re_ttlmarker++ ){
        this.result_ttlmarker =  Number(this.result_ttlmarker) + Number(this.rowexport[re_ttlmarker].ttl_marker)
        
        
      }
      this.keep_ttlmarker.push({
          result_ttlmarker:this.result_ttlmarker,
        })
    this.sum_lenth_yd = 0
    this.sum_all_lenth_yd = 0
    for(var ax = 0; ax<this.rowexport.length; ax++){
       this.sum_lenth_yd = this.rowexport[ax].length_ydb * this.rowexport[ax].ttl_marker
       this.sum_all_lenth_yd = Number(this.sum_all_lenth_yd) + Number(this.sum_lenth_yd)
    }
     
    this.total_lenth_yd_j = this.sum_all_lenth_yd  / this.result_ttlmarker
        worksheet.getCell("J"+ty).value = Number(this.total_lenth_yd_j);
        worksheet.getCell("J"+ty).numFmt = "0.00"
      worksheet.getCell("J"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.keep_avg_marker_length.push({
        value:this.total_lenth_yd_j,
        org:this.multiple[x]
      })

      this.keep_sum_x.push({
        value:this.sum_all_lenth_yd
      })
      this.keep_sum_ttl_x.push({
          value:this.result_ttlmarker
      })

      this.result_plug12 = 0
      for(var re_plug1_2 = 0; re_plug1_2<resp.data.data.length; re_plug1_2++ ){
        this.result_plug12 =  Number(this.result_plug12) + Number(this.rowexport[re_plug1_2].plug12)
        
      }
     
     this.result_plug12 = this.result_plug12/resp.data.data.length 
     this.keep_plug12_inch.push({
          result_plug12_inch:this.result_plug12,
        })

        worksheet.getCell("K"+ty).value = Number(this.result_ttlmarker)
        worksheet.getCell("K"+ty).numFmt = '0.00';
      worksheet.getCell("K"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 

      
      this.result_perplug12 = 0
      for(var re_perplug12 = 0; re_perplug12<resp.data.data.length; re_perplug12++ ){
        this.result_perplug12 =  Number(this.result_perplug12) + Number(this.rowexport[re_perplug12].widthloss)
        
      }
      this.keep_perplug12.push({
          result_perplug12:this.result_perplug12,
        })

     
     this.result_perplug12 = this.result_perplug12/resp.data.data.length 

        worksheet.getCell("L"+ty).value = "";
      worksheet.getCell("L"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
      
      for(var replug12_inch = 0; replug12_inch<resp.data.data.length; replug12_inch++){
      this.result_plug12_inch = (this.rowexport[replug12_inch].endless_length_yd*36)/(this.rowexport[replug12_inch].ttl_marker/this.rowexport[replug12_inch].length_ydb)
    
      this.end_loss_plug12.push(this.result_plug12_inch)
        
      }
        this.result_plug12_avg = 0;
        for(var replug12_avg = 0; replug12_avg<this.end_loss_plug12.length; replug12_avg++){
        this.result_plug12_avg= Number(this.result_plug12_avg) + Number(this.end_loss_plug12[replug12_avg])
      
      }  
     
        this.all_result_avg = ((this.result_plug12_avg * this.result_ttlmarker) / this.result_ttlmarker)
      
      
        
    
    this.result_length_yd = 0
      for(var re_length_yd = 0; re_length_yd<resp.data.data.length; re_length_yd++ ){
        this.result_length_yd =  Number(this.result_length_yd) + Number(this.rowexport[re_length_yd].endless_length_yd)
        
      }
     this.keep_avg_length_yd.push({
       all_result_avg_length:this.result_length_yd
     })

        worksheet.getCell("N"+ty).value = Number(this.result_length_yd)
      worksheet.getCell("N"+ty).numFmt = '0.00';
      worksheet.getCell("N"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       this.result_avg_per = 0;
      this.result_avg_per = this.result_length_yd / this.result_ttlmarker*100
      
      this.keep_avg_end.push({
       result_avg_per: this.result_avg_per
      })
      
      worksheet.getCell("O"+ty).value = this.result_avg_per/100
      worksheet.getCell("O"+ty).numFmt = '0.00%';
      worksheet.getCell("O"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_splice = 0
      for(var re_splice = 0; re_splice<resp.data.data.length; re_splice++ ){
        this.result_splice =  Number(this.result_splice) + Number(this.rowexport[re_splice].spliceloss)
        
      }
    
     this.keep_splice_Loss.push({
          result_splice:this.result_splice
     })


        worksheet.getCell("P"+ty).value = Number(this.result_splice)
      worksheet.getCell("P"+ty).numFmt = '0.00';
      worksheet.getCell("P"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_splicelossper = 0;
      this.result_splicelossper = this.result_splice / this.result_ttlmarker*100
      worksheet.getCell("Q"+ty).value = this.result_splicelossper / 100
      worksheet.getCell("Q"+ty).numFmt = '0.00%';
      worksheet.getCell("Q"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


       this.result_cut_out_loss = 0
      for(var re_cut_out = 0; re_cut_out<resp.data.data.length; re_cut_out++ ){
        this.result_cut_out_loss =  Number(this.result_cut_out_loss) + Number(this.rowexport[re_cut_out].totalcutout)
        
      }
   
     this.keep_cut_out_loss.push({
       result_cut_out_loss: this.result_cut_out_loss
     })

        worksheet.getCell("R"+ty).value = Number(this.result_cut_out_loss)
      worksheet.getCell("R"+ty).numFmt = '0.00';
      worksheet.getCell("R"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
      this.result_cut_out_loss_per = 0;
      this.result_cut_out_loss_per = this.result_cut_out_loss / this.result_ttlmarker*100
      worksheet.getCell("S"+ty).value = this.result_cut_out_loss_per/100
      worksheet.getCell("S"+ty).numFmt = '0.00%';
      worksheet.getCell("S"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_remnants = 0
      for(var re_remnants = 0; re_remnants<resp.data.data.length; re_remnants++ ){
        this.result_remnants =  Number(this.result_remnants) + Number(this.rowexport[re_remnants].rement)
        
      }
     this.sum_remnants = this.result_remnants
     
     this.keep_remnants.push({
       result_remnants:this.result_remnants
     })

       worksheet.getCell("T"+ty).value = Number(this.result_remnants)
      worksheet.getCell("T"+ty).numFmt = '0.00';
      worksheet.getCell("T"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_remnants_per = 0;
      this.result_remnants_per = this.result_remnants / this.result_ttlmarker*100
      worksheet.getCell("U"+ty).value = this.result_remnants_per/100
      worksheet.getCell("U"+ty).numFmt = '0.00%';
      worksheet.getCell("U"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


      this.result_total_length = 0
      for(var re_total_length = 0; re_total_length<resp.data.data.length; re_total_length++ ){
      this.result_total_length =  Number(this.result_total_length) + Number(this.rowexport[re_total_length].totallenthloss)
        
      }
     
     this.keep_total_loss.push({
       result_total_length:this.result_total_length
     }) 

        worksheet.getCell("V"+ty).value = Number(this.result_total_length)
      worksheet.getCell("V"+ty).numFmt = '0.00';
      worksheet.getCell("V"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };








       //B---------------------------------------------------------------------
    

        //C------------------------------------------------------------------------
        worksheet.getCell("B"+a).value = "S/O no.";
      worksheet.mergeCells("B"+a +":" + "B"+b);
      worksheet.getCell("B"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("B"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "B" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].gm_no_short);
     
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

        //D------------------------------------------------------------------------
        worksheet.getCell("C"+a).value = "Customer";
      worksheet.mergeCells("C"+a +":" + "C"+b);
      worksheet.getCell("C"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("C"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "C" + ab
      
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
      worksheet.getCell("D"+a).value = "GM08 no.";
      worksheet.mergeCells("D"+a +":" + "D"+b);
      worksheet.getCell("D"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("D"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "D" + ab
      
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
      worksheet.getCell("E"+a).value = this.multiple_rec;
     worksheet.mergeCells("E"+a +":" + "E"+b);
      worksheet.getCell("E"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("E"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "E" + ab
      
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
        worksheet.getCell("F"+a).value = "Type"
     worksheet.mergeCells("F"+a +":" + "F"+b);
      worksheet.getCell("F"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("F"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "F" + ab
      
            var z = ab-a-2;

            
       worksheet.getCell(y).value = Number(this.rowexport[z].f_type);
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("G"+b).value = "Width";
      worksheet.getCell("G"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("G"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "G" + ab
      
            var z = ab-a-2;
      this.obs_width_last = parseFloat(this.rowexport[z].obs_width).toFixed(2)
       worksheet.getCell(y).value = Number(this.obs_width_last);
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("H"+b).value = "(inch)";
      worksheet.getCell("H"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "H" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].width_inch)
       worksheet.getCell(y).numFmt = '0.00'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
      //I-----------------------------------------------------------

          worksheet.getCell("J"+a).value = "Length";
      worksheet.getCell("J"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("J"+b).value = "(yd)";
      worksheet.getCell("J"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
      this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "J" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value =  Number(this.rowexport[z].length_ydb)
       worksheet.getCell(y).numFmt = '0.00'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       } 
        worksheet.getCell("I"+a).value = "Marker VS";
      worksheet.getCell("I"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("I"+b).value = "obs.width";
      worksheet.getCell("I"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
      this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "I" + ab
      
            var z = ab-a-2;

          this.mark_obs = Number(this.rowexport[z].obs_width)-Number(this.rowexport[z].width_inch) 
       worksheet.getCell(y).value = Number(this.mark_obs)
       worksheet.getCell(y).numFmt = '0.00'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }


       //J-----------------------------------------------------------
         worksheet.getCell("K"+a).value = "Ttl Marker";
      worksheet.getCell("K"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("K"+b).value = "Length(yd)";
      worksheet.getCell("K"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "K" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].ttl_marker)
       worksheet.getCell(y).numFmt = '0.00'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

//K-------------------------------------------------------------------------------
         worksheet.getCell("L"+a).value = "Plug1+2";
      worksheet.getCell("L"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("L"+b).value = "(inch)";
      worksheet.getCell("L"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
         this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "L" + ab
      
            var z = ab-a-2;
      worksheet.getCell(y).value = Number(this.rowexport[z].plug12)
      worksheet.getCell(y).numFmt = '0.00'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

//L------------------------------------------------------------------------------------
         worksheet.getCell("M"+a).value = "%";
      worksheet.getCell("M"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("M"+b).value = 1.50/100;
       worksheet.getCell("M"+b).numFmt = '0.00%'; 
      worksheet.getCell("M"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "M" + ab
      
            var z = ab-a-2;
      worksheet.getCell(y).value = this.rowexport[z].widthloss/100
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
      worksheet.getCell("M"+ty).value = "";
      worksheet.getCell("M"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

    //N----------------------------------------------------------
       worksheet.getCell("N"+a).value = "Length";
      worksheet.getCell("N"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("N"+b).value = "(yd)";
      worksheet.getCell("N"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
        this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "N" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].endless_length_yd)
       worksheet.getCell(y).numFmt = '0.00'; 
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 

     this.check_value = ""
      if(this.multiple[x] == "Woven"){
       this.check_value =  this.end_loss_value[0].p1;
      }else if(this.multiple[x] == "Light-Middle Weight"){
       this.check_value =  this.end_loss_value[0].p2;
      }else if(this.multiple[x] == "fleece"){
        this.check_value =  this.end_loss_value[0].p3;
      }else{
        this.check_value = this.end_loss_value[0].p4;
      }

    worksheet.getCell("O"+b).value = this.check_value / 100;
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
       worksheet.getCell(y).value = this.rowexport[z].avg_end/100
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("P"+b).value = "(yd)";
      worksheet.getCell("P"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "P" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].spliceloss)
       worksheet.getCell(y).numFmt = '0.00'; 
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
      if(this.multiple[x] == "Woven"){
       worksheet.getCell("Q"+b).value = this.splice_value[0].p1 / 100;
      }else if(this.multiple[x] == "Light-Middle Weight"){
       worksheet.getCell("Q"+b).value = this.splice_value[0].p2 / 100;
      }else if(this.multiple[x] == "Fleece"){
        worksheet.getCell("Q"+b).value = this.splice_value[0].p3 / 100;
      }else{
        worksheet.getCell("Q"+b).value = this.splice_value[0].p4 / 100;
      }
 
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
       worksheet.getCell(y).value = this.rowexport[z].splicelossper/100
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("R"+b).value = "(yd)";
      worksheet.getCell("R"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
          this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "R" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].totalcutout)
       worksheet.getCell(y).numFmt = '0.00'; 
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 

       if(this.multiple[x] == "Woven"){
       worksheet.getCell("S"+b).value = this.cut_out_value[0].p1 / 100;
      }else if(this.multiple[x] == "Light-Middle Weight"){
       worksheet.getCell("S"+b).value = this.cut_out_value[0].p2 / 100;
      }else if(this.multiple[x] == "Fleece"){
        worksheet.getCell("S"+b).value = this.cut_out_value[0].p3 / 100;
      }else{
        worksheet.getCell("S"+b).value = this.cut_out_value[0].p4 / 100;
      }
       worksheet.getCell("S"+b).numFmt = '0.00%'; 
      worksheet.getCell("S"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("T"+b).value = "(yd)";
      worksheet.getCell("T"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "T" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].rement)
       worksheet.getCell(y).numFmt = '0.00'; 
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       if(this.multiple[x] == "Woven"){
       worksheet.getCell("U"+b).value = this.remnants_loss_value[0].p1 / 100;
      }else if(this.multiple[x] == "Light-Middle Weight"){
       worksheet.getCell("U"+b).value = this.remnants_loss_value[0].p2 / 100;
      }else if(this.multiple[x] == "Fleece"){
        worksheet.getCell("U"+b).value = this.remnants_loss_value[0].p3 / 100;
      }else{
        worksheet.getCell("U"+b).value = this.remnants_loss_value[0].p4 / 100;
      }
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
       worksheet.getCell(y).value = this.rowexport[z].rement_per/100
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
        if(this.multiple[x] == "Woven"){
       worksheet.getCell("V"+b).value = this.remnants_loss_value[0].p1 / 100;
      }else if(this.multiple[x] == "Light-Middle Weight"){
       worksheet.getCell("V"+b).value = this.remnants_loss_value[0].p2 / 100;
      }else if(this.multiple[x] == "fleece"){
        worksheet.getCell("V"+b).value = this.remnants_loss_value[0].p3 / 100;
      }else{
        worksheet.getCell("V"+b).value = this.remnants_loss_value[0].p4 / 100;
      }

       worksheet.getCell("V"+b).value = "(yd)";
      worksheet.getCell("V"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "V" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].totallenthloss)
       worksheet.getCell(y).numFmt = '0.00'; 
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       if(this.multiple[x] == "Woven"){
       worksheet.getCell("W"+b).value = this.total_woven / 100;
      }else if(this.multiple[x] == "Light-Middle Weight"){
       worksheet.getCell("W"+b).value = this.total_light / 100;
      }else if(this.multiple[x] == "Fleece"){
        worksheet.getCell("W"+b).value = this.total_fleece / 100;
      }else{
        worksheet.getCell("W"+b).value = this.total_special / 100;
      }
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
       worksheet.getCell(y).value = this.rowexport[z].totallenthlossper/100
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
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
    
       worksheet.getCell("A"+ty).value = "Total Type :" + this.multiple_rec;
      worksheet.mergeCells("A"+ty +":" + "I"+ty);
      worksheet.getCell("A"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
        this.result_ttlmarker = 0
      for(var re_ttlmarker = 0; re_ttlmarker<resp.data.data.length; re_ttlmarker++ ){
        this.result_ttlmarker =  Number(this.result_ttlmarker) + Number(this.rowexport[re_ttlmarker].ttl_marker)
      
      }
        this.keep_ttlmarker.push({
          result_ttlmarker:this.result_ttlmarker,
        })
 this.sum_lenth_yd = 0
    this.sum_all_lenth_yd = 0
    for(var ax = 0; ax<this.rowexport.length; ax++){
       this.sum_lenth_yd = this.rowexport[ax].length_ydb * this.rowexport[ax].ttl_marker
       this.sum_all_lenth_yd = Number(this.sum_all_lenth_yd) + Number(this.sum_lenth_yd)
    }
     console.log(this.sum_all_lenth_yd )
    this.total_lenth_yd_j = this.sum_all_lenth_yd  / this.result_ttlmarker
        worksheet.getCell("J"+ty).value = Number(this.total_lenth_yd_j);
        worksheet.getCell("J"+ty).numFmt = "0.00"
      worksheet.getCell("J"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };  

      this.keep_avg_marker_length.push({
        value:this.total_lenth_yd_j,
        org:this.multiple[x]
      })

      this.keep_sum_x.push({
        value:this.sum_all_lenth_yd
      })
      this.keep_sum_ttl_x.push({
          value:this.result_ttlmarker
      })
      
      this.result_plug12 = 0
      for(var re_plug1_2 = 0; re_plug1_2<resp.data.data.length; re_plug1_2++ ){
        this.result_plug12 =  Number(this.result_plug12) + Number(this.rowexport[re_plug1_2].plug12)
        
      }
     
     this.result_plug12 = this.result_plug12/resp.data.data.length 
     this.keep_plug12_inch.push({
          result_plug12_inch:this.result_plug12,
        })


        worksheet.getCell("K"+ty).value = Number(this.result_ttlmarker)
        worksheet.getCell("K"+ty).numFmt = '0.00'; 
      worksheet.getCell("K"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 

       this.result_perplug12 = 0
      for(var re_perplug12 = 0; re_perplug12<resp.data.data.length; re_perplug12++ ){
        this.result_perplug12 =  Number(this.result_perplug12) + Number(this.rowexport[re_perplug12].widthloss)
        
      }
   this.keep_perplug12.push({
          result_perplug12:this.result_perplug12,
        })
     this.result_perplug12_l = this.result_perplug12/resp.data.data.length 
     

        worksheet.getCell("L"+ty).value = "";
      worksheet.getCell("L"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       for(var replug12_inch = 0; replug12_inch<resp.data.data.length; replug12_inch++){
      this.result_plug12_inch = (this.rowexport[replug12_inch].endless_length_yd*36)/(this.rowexport[replug12_inch].ttl_marker/this.rowexport[replug12_inch].length_ydb)
    
      this.end_loss_plug12.push(this.result_plug12_inch)
        
      }
        this.result_plug12_avg = 0;
        for(var replug12_avg = 0; replug12_avg<this.end_loss_plug12.length; replug12_avg++){
        this.result_plug12_avg= Number(this.result_plug12_avg) + Number(this.end_loss_plug12[replug12_avg])
      
      }  
     
        this.all_result_avg = ((this.result_plug12_avg * this.result_ttlmarker) / this.result_ttlmarker)
        this.keep_avg_loss_plug12.push({
          all_result_avg:this.all_result_avg
        })    

        
    
    

      this.result_length_yd = 0
      for(var re_length_yd = 0; re_length_yd<resp.data.data.length; re_length_yd++ ){
        this.result_length_yd =  Number(this.result_length_yd) + Number(this.rowexport[re_length_yd].endless_length_yd)
        
      }
      this.keep_avg_length_yd.push({
       all_result_avg_length:this.result_length_yd
     })
     

        worksheet.getCell("N"+ty).value = Number(this.result_length_yd)
        worksheet.getCell("N"+ty).numFmt = '0.00'; 
      worksheet.getCell("N"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


      
       this.result_avg_per = 0;
      this.result_avg_per = this.result_length_yd / this.result_ttlmarker*100
      
      this.keep_avg_end.push({
       result_avg_per: this.result_avg_per
      })
      
      worksheet.getCell("O"+ty).value = this.result_avg_per/100
      worksheet.getCell("O"+ty).numFmt = '0.00%';
      worksheet.getCell("O"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_splice = 0
      for(var re_splice = 0; re_splice<resp.data.data.length; re_splice++ ){
        this.result_splice =  Number(this.result_splice) + Number(this.rowexport[re_splice].spliceloss)
        
      }
     
    
     this.keep_splice_Loss.push({
          result_splice:this.result_splice
     })

      worksheet.getCell("P"+ty).value = Number(this.result_splice)
      worksheet.getCell("P"+ty).numFmt = '0.00';
      worksheet.getCell("P"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_splicelossper = 0;
      this.result_splicelossper = this.result_splice / this.result_ttlmarker*100
         worksheet.getCell("Q"+ty).value = this.result_splicelossper / 100
      worksheet.getCell("Q"+ty).numFmt = '0.00%';
      worksheet.getCell("Q"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


       this.result_cut_out_loss = 0
      for(var re_cut_out = 0; re_cut_out<resp.data.data.length; re_cut_out++ ){
        this.result_cut_out_loss =  Number(this.result_cut_out_loss) + Number(this.rowexport[re_cut_out].totalcutout)
        
      }
     
   
     this.keep_cut_out_loss.push({
       result_cut_out_loss: this.result_cut_out_loss
     })

        worksheet.getCell("R"+ty).value = Number(this.result_cut_out_loss)
      worksheet.getCell("R"+ty).numFmt = '0.00';
      worksheet.getCell("R"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
      this.result_cut_out_loss_per = 0;
      this.result_cut_out_loss_per = this.result_cut_out_loss / this.result_ttlmarker*100
      worksheet.getCell("S"+ty).value = this.result_cut_out_loss_per/100
      worksheet.getCell("S"+ty).numFmt = '0.00%';
      worksheet.getCell("S"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_remnants = 0
      for(var re_remnants = 0; re_remnants<resp.data.data.length; re_remnants++ ){
        this.result_remnants =  Number(this.result_remnants) + Number(this.rowexport[re_remnants].rement)
        
      }
     
     
     this.keep_remnants.push({
       result_remnants:this.result_remnants
     })

      worksheet.getCell("T"+ty).value = Number(this.result_remnants)
      worksheet.getCell("T"+ty).numFmt = '0.00';
      worksheet.getCell("T"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_remnants_per = 0;
      this.result_remnants_per = this.result_remnants / this.result_ttlmarker*100
      worksheet.getCell("U"+ty).value = this.result_remnants_per/100
      worksheet.getCell("U"+ty).numFmt = '0.00%';
      worksheet.getCell("U"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


      this.result_total_length = 0
      for(var re_total_length = 0; re_total_length<resp.data.data.length; re_total_length++ ){
      this.result_total_length =  Number(this.result_total_length) + Number(this.rowexport[re_total_length].totallenthloss)
        
      }
     
    
     this.keep_total_loss.push({
       result_total_length:this.result_total_length
     }) 

      worksheet.getCell("V"+ty).value = Number(this.result_total_length)
      worksheet.getCell("V"+ty).numFmt = '0.00';
      worksheet.getCell("V"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };



       //B---------------------------------------------------------------------

     

       //C---------------------------------------------------------------------

       worksheet.getCell("B"+a).value = "S/O no.";
      worksheet.mergeCells("B"+a +":" + "B"+b);
      worksheet.getCell("B"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("B"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "B" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].gm_no_short);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       
       //D---------------------------------------------------------------------

      worksheet.getCell("C"+a).value = "Customer";
      worksheet.mergeCells("C"+a +":" + "C"+b);
      worksheet.getCell("C"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("C"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "C" + ab
      
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
       
     worksheet.getCell("D"+a).value = "GM08 no.";
      worksheet.mergeCells("D"+a +":" + "D"+b);
      worksheet.getCell("D"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("D"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "D" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_no
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
    
      worksheet.getCell("E"+a).value = this.multiple_rec;
     worksheet.mergeCells("E"+a +":" + "E"+b);
      worksheet.getCell("E"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("E"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "E" + ab
      
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
        worksheet.getCell("F"+a).value = "Type"
     worksheet.mergeCells("F"+a +":" + "F"+b);
      worksheet.getCell("F"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("F"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "F" + ab
      
            var z = ab-a-2;

            
       worksheet.getCell(y).value = Number(this.rowexport[z].f_type);
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("G"+b).value = "Width";
      worksheet.getCell("G"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("G"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "G" + ab
      
            var z = ab-a-2;
       this.obs_width_last = parseFloat(this.rowexport[z].obs_width).toFixed(2)
       worksheet.getCell(y).value = Number(this.obs_width_last);
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("H"+b).value = "(inch)";
      worksheet.getCell("H"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "H" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].width_inch)
       worksheet.getCell(y).numFmt = '0.00'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //I--------------------------------------------------------------
          worksheet.getCell("J"+a).value = "Length";
      worksheet.getCell("J"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("J"+b).value = "(yd)";
      worksheet.getCell("J"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
      this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "J" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].length_ydb)
       worksheet.getCell(y).numFmt = '0.00'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       } 
        worksheet.getCell("I"+a).value = "Marker VS";
      worksheet.getCell("I"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("I"+b).value = "obs.width";
      worksheet.getCell("I"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
      this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "I" + ab
      
            var z = ab-a-2;

          this.mark_obs = Number(this.rowexport[z].obs_width)-Number(this.rowexport[z].width_inch) 
       worksheet.getCell(y).value = Number(this.mark_obs)
       worksheet.getCell(y).numFmt = '0.00'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
       //J-----------------------------------------------------------
         worksheet.getCell("K"+a).value = "Ttl Marker";
      worksheet.getCell("K"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("K"+b).value = "Length(yd)";
      worksheet.getCell("K"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "K" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].ttl_marker)
       worksheet.getCell(y).numFmt = '0.00'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
       }; 
       }

        //K-------------------------------------------------------------------------------
         worksheet.getCell("L"+a).value = "Plug1+2";
      worksheet.getCell("L"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("L"+b).value = "(inch)";
      worksheet.getCell("L"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
         
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "L" + ab
      
            var z = ab-a-2;
      worksheet.getCell(y).value = Number(this.rowexport[z].plug12)
      worksheet.getCell(y).numFmt = '0.00'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
       }; 
       }

       //L------------------------------------------------------------------
            worksheet.getCell("M"+a).value = "%";
      worksheet.getCell("M"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("M"+b).value = 1.50/100
       worksheet.getCell("M"+b).numFmt = '0.00%'; 
      worksheet.getCell("M"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

          this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "M" + ab
      
            var z = ab-a-2;
      worksheet.getCell(y).value = this.rowexport[z].widthloss /100
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
      worksheet.getCell("M"+ty).value = "";
      worksheet.getCell("M"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

 //N----------------------------------------------------------
       worksheet.getCell("N"+a).value = "Length";
      worksheet.getCell("N"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("N"+b).value = "(yd)";
      worksheet.getCell("N"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "N" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].endless_length_yd)
       worksheet.getCell(y).numFmt = '0.00'; 
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.check_value = ""
      if(this.multiple[x] == "Woven"){
       this.check_value =  this.end_loss_value[0].p1;
      }else if(this.multiple[x] == "Light-Middle Weight"){
       this.check_value =  this.end_loss_value[0].p2;
      }else if(this.multiple[x] == "fleece"){
        this.check_value =  this.end_loss_value[0].p3;
      }else{
        this.check_value = this.end_loss_value[0].p4;
      }

    worksheet.getCell("O"+b).value = this.check_value / 100;
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
       worksheet.getCell(y).value = this.rowexport[z].avg_end/100
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("P"+b).value = "(yd)";
      worksheet.getCell("P"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "P" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].spliceloss)
       worksheet.getCell(y).numFmt = '0.00'; 
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
      if(this.multiple[x] == "Woven"){
       worksheet.getCell("Q"+b).value = this.splice_value[0].p1 / 100;
      }else if(this.multiple[x] == "Light-Middle Weight"){
       worksheet.getCell("Q"+b).value = this.splice_value[0].p2 / 100;
      }else if(this.multiple[x] == "Fleece"){
        worksheet.getCell("Q"+b).value = this.splice_value[0].p3 / 100;
      }else{
        worksheet.getCell("Q"+b).value = this.splice_value[0].p4 / 100;
      }
 
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
       worksheet.getCell(y).value = this.rowexport[z].splicelossper/100
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("R"+b).value = "(yd)";
      worksheet.getCell("R"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
          this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "R" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].totalcutout)
       worksheet.getCell(y).numFmt = '0.00'; 
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
      if(this.multiple[x] == "Woven"){
       worksheet.getCell("S"+b).value = this.cut_out_value[0].p1 / 100;
      }else if(this.multiple[x] == "Light-Middle Weight"){
       worksheet.getCell("S"+b).value = this.cut_out_value[0].p2 / 100;
      }else if(this.multiple[x] == "Fleece"){
        worksheet.getCell("S"+b).value = this.cut_out_value[0].p3 / 100;
      }else{
        worksheet.getCell("S"+b).value = this.cut_out_value[0].p4 / 100;
      }
       worksheet.getCell("S"+b).numFmt = '0.00%'; 
      worksheet.getCell("S"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("T"+b).value = "(yd)";
      worksheet.getCell("T"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
        this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "T" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].rement)
       worksheet.getCell(y).numFmt = '0.00'; 
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       if(this.multiple[x] == "Woven"){
       worksheet.getCell("U"+b).value = this.remnants_loss_value[0].p1 / 100;
      }else if(this.multiple[x] == "Light-Middle Weight"){
       worksheet.getCell("U"+b).value = this.remnants_loss_value[0].p2 / 100;
      }else if(this.multiple[x] == "Fleece"){
        worksheet.getCell("U"+b).value = this.remnants_loss_value[0].p3 / 100;
      }else{
        worksheet.getCell("U"+b).value = this.remnants_loss_value[0].p4 / 100;
      }
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
       worksheet.getCell(y).value = this.rowexport[z].rement_per/100
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("V"+b).value = "(yd)";
      worksheet.getCell("V"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "V" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].totallenthloss)
       worksheet.getCell(y).numFmt = '0.00'; 
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
      if(this.multiple[x] == "Woven"){
       worksheet.getCell("W"+b).value = this.total_woven / 100;
      }else if(this.multiple[x] == "Light-Middle Weight"){
       worksheet.getCell("W"+b).value = this.total_light / 100;
      }else if(this.multiple[x] == "Fleece"){
        worksheet.getCell("W"+b).value = this.total_fleece / 100;
      }else{
        worksheet.getCell("W"+b).value = this.total_special / 100;
      }
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
       worksheet.getCell(y).value = this.rowexport[z].totallenthlossper/100
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
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
    
       worksheet.getCell("A"+ty).value = "Total Type :" + this.multiple_rec;
      worksheet.mergeCells("A"+ty +":" + "I"+ty);
      worksheet.getCell("A"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
      
    this.sum_lenth_yd = 0
    this.sum_all_lenth_yd = 0
    for(var ax = 0; ax<this.rowexport.length; ax++){
       this.sum_lenth_yd = this.rowexport[ax].length_ydb * this.rowexport[ax].ttl_marker
       this.sum_all_lenth_yd = Number(this.sum_all_lenth_yd) + Number(this.sum_lenth_yd)
    }
     console.log(this.sum_all_lenth_yd )
    this.total_lenth_yd_j = this.sum_all_lenth_yd  / this.result_ttlmarker
        worksheet.getCell("J"+ty).value = Number(this.total_lenth_yd_j);
        worksheet.getCell("J"+ty).numFmt = "0.00"
      worksheet.getCell("J"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.keep_avg_marker_length.push({
        value:this.total_lenth_yd_j,
        org:this.multiple[x]
      })

      this.keep_sum_x.push({
        value:this.sum_all_lenth_yd
      })
      this.keep_sum_ttl_x.push({
          value:this.result_ttlmarker
      })

     this.result_plug12 = 0
      for(var re_plug1_2 = 0; re_plug1_2<resp.data.data.length; re_plug1_2++ ){
        this.result_plug12 =  Number(this.result_plug12) + Number(this.rowexport[re_plug1_2].plug12)
        
      }
     
     this.result_plug12 = this.result_plug12/resp.data.data.length 
     this.keep_plug12_inch.push({
          result_plug12_inch:this.result_plug12,
        })


 this.result_ttlmarker = 0
      for(var re_ttlmarker = 0; re_ttlmarker<resp.data.data.length; re_ttlmarker++ ){
        this.result_ttlmarker =  Number(this.result_ttlmarker) + Number(this.rowexport[re_ttlmarker].ttl_marker)
        
      }
      this.keep_ttlmarker.push({
          result_ttlmarker:this.result_ttlmarker,
        })

          worksheet.getCell("K"+ty).value = Number(this.result_ttlmarker)
      worksheet.getCell("K"+ty).numFmt = '0.00';
      worksheet.getCell("K"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 

       this.result_perplug12 = 0
      for(var re_perplug12 = 0; re_perplug12<resp.data.data.length; re_perplug12++ ){
        this.result_perplug12 =  Number(this.result_perplug12) + Number(this.rowexport[re_perplug12].widthloss)
        
      }
       this.keep_perplug12.push({
          result_perplug12:this.result_perplug12,
        })

     
     this.result_perplug12 = this.result_perplug12/resp.data.data.length 

        worksheet.getCell("L"+ty).value = "";
      worksheet.getCell("L"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       for(var replug12_inch = 0; replug12_inch<resp.data.data.length; replug12_inch++){
      this.result_plug12_inch = (this.rowexport[replug12_inch].endless_length_yd*36)/(this.rowexport[replug12_inch].ttl_marker/this.rowexport[replug12_inch].length_ydb)
    
      this.end_loss_plug12.push(this.result_plug12_inch)
        
      }
        this.result_plug12_avg = 0;
        for(var replug12_avg = 0; replug12_avg<this.end_loss_plug12.length; replug12_avg++){
        this.result_plug12_avg= Number(this.result_plug12_avg) + Number(this.end_loss_plug12[replug12_avg])
      
      }  
     
        this.all_result_avg = ((this.result_plug12_avg * this.result_ttlmarker) / this.result_ttlmarker)
      
      
     
this.result_length_yd = 0
      for(var re_length_yd = 0; re_length_yd<resp.data.data.length; re_length_yd++ ){
        this.result_length_yd =  Number(this.result_length_yd) + Number(this.rowexport[re_length_yd].endless_length_yd)
        
      }
     this.keep_avg_length_yd.push({
       all_result_avg_length:this.result_length_yd
     })
     

        worksheet.getCell("N"+ty).value = Number(this.result_length_yd)
        worksheet.getCell("N"+ty).numFmt = '0.00'; 
      worksheet.getCell("N"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
       this.result_avg_per = 0;
      this.result_avg_per = this.result_length_yd / this.result_ttlmarker*100
      
      this.keep_avg_end.push({
       result_avg_per: this.result_avg_per
      })
      
      worksheet.getCell("O"+ty).value = this.result_avg_per/100
      worksheet.getCell("O"+ty).numFmt = '0.00%';
      worksheet.getCell("O"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_splice = 0
      for(var re_splice = 0; re_splice<resp.data.data.length; re_splice++ ){
        this.result_splice =  Number(this.result_splice) + Number(this.rowexport[re_splice].spliceloss)
        
      }
     
     
     this.keep_splice_Loss.push({
          result_splice:this.result_splice
     })


      worksheet.getCell("P"+ty).value = Number(this.result_splice)
      worksheet.getCell("P"+ty).numFmt = '0.00';
      worksheet.getCell("P"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_splicelossper = 0;
      this.result_splicelossper = this.result_splice / this.result_ttlmarker*100
         worksheet.getCell("Q"+ty).value = this.result_splicelossper / 100
      worksheet.getCell("Q"+ty).numFmt = '0.00%';
      worksheet.getCell("Q"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


       this.result_cut_out_loss = 0
      for(var re_cut_out = 0; re_cut_out<resp.data.data.length; re_cut_out++ ){
        this.result_cut_out_loss =  Number(this.result_cut_out_loss) + Number(this.rowexport[re_cut_out].totalcutout)
        
      }
     

     this.keep_cut_out_loss.push({
       result_cut_out_loss: this.result_cut_out_loss
     })

        worksheet.getCell("R"+ty).value = Number(this.result_cut_out_loss)
      worksheet.getCell("R"+ty).numFmt = '0.00';
      worksheet.getCell("R"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
      this.result_cut_out_loss_per = 0;
      this.result_cut_out_loss_per = this.result_cut_out_loss / this.result_ttlmarker*100
      worksheet.getCell("S"+ty).value = this.result_cut_out_loss_per/100
      worksheet.getCell("S"+ty).numFmt = '0.00%';
      worksheet.getCell("S"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_remnants = 0
      for(var re_remnants = 0; re_remnants<resp.data.data.length; re_remnants++ ){
        this.result_remnants =  Number(this.result_remnants) + Number(this.rowexport[re_remnants].rement)
        
      }
     
     
     this.keep_remnants.push({
       result_remnants:this.result_remnants
     })

      worksheet.getCell("T"+ty).value = Number(this.result_remnants)
      worksheet.getCell("T"+ty).numFmt = '0.00';
      worksheet.getCell("T"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_remnants_per = 0;
      this.result_remnants_per = this.result_remnants / this.result_ttlmarker*100
      worksheet.getCell("U"+ty).value = this.result_remnants_per/100
      worksheet.getCell("U"+ty).numFmt = '0.00%';
      worksheet.getCell("U"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


      this.result_total_length = 0
      for(var re_total_length = 0; re_total_length<resp.data.data.length; re_total_length++ ){
      this.result_total_length =  Number(this.result_total_length) + Number(this.rowexport[re_total_length].totallenthloss)
        
      }
     
     
     this.keep_total_loss.push({
       result_total_length:this.result_total_length
     }) 

       worksheet.getCell("V"+ty).value = Number(this.result_total_length)
      worksheet.getCell("V"+ty).numFmt = '0.00';
      worksheet.getCell("V"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       //B----------------------------------------------------------------
    

       //C---------------------------------------------------------------------

       worksheet.getCell("B"+a).value = "S/O no.";
      worksheet.mergeCells("B"+a +":" + "B"+b);
      worksheet.getCell("B"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("B"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "B" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].gm_no_short);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

           //D---------------------------------------------------------------------

        worksheet.getCell("C"+a).value = "Customer";
      worksheet.mergeCells("C"+a +":" + "C"+b);
      worksheet.getCell("C"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("C"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "C" + ab
      
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
      worksheet.getCell("D"+a).value = "GM08 no.";
      worksheet.mergeCells("D"+a +":" + "D"+b);
      worksheet.getCell("D"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("D"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "D" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_no
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
     worksheet.getCell("E"+a).value = this.multiple_rec;
     worksheet.mergeCells("E"+a +":" + "E"+b);
      worksheet.getCell("E"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("E"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "E" + ab
      
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
        worksheet.getCell("F"+a).value = "Type"
     worksheet.mergeCells("F"+a +":" + "F"+b);
      worksheet.getCell("F"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("F"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "F" + ab
      
            var z = ab-a-2;

            
       worksheet.getCell(y).value = Number(this.rowexport[z].f_type);
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("G"+b).value = "Width";
      worksheet.getCell("G"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("G"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "G" + ab
      
            var z = ab-a-2;
       this.obs_width_last = parseFloat(this.rowexport[z].obs_width).toFixed(2)
       worksheet.getCell(y).value = Number(this.obs_width_last);
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("H"+b).value = "(inch)";
      worksheet.getCell("H"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "H" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].width_inch)
       worksheet.getCell(y).numFmt = '0.00'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //I--------------------------------------------------------------
          worksheet.getCell("J"+a).value = "Length";
      worksheet.getCell("J"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("J"+b).value = "(yd)";
      worksheet.getCell("J"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
      this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "J" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].length_ydb)
       worksheet.getCell(y).numFmt = '0.00'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       } 
        worksheet.getCell("I"+a).value = "Marker VS";
      worksheet.getCell("I"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("I"+b).value = "obs.width";
      worksheet.getCell("I"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
      this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "I" + ab
      
            var z = ab-a-2;

          this.mark_obs = Number(this.rowexport[z].obs_width)-Number(this.rowexport[z].width_inch) 
       worksheet.getCell(y).value = Number(this.mark_obs)
       worksheet.getCell(y).numFmt = '0.00'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

         //J-----------------------------------------------------------
         worksheet.getCell("K"+a).value = "Ttl Marker";
      worksheet.getCell("K"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("K"+b).value = "Length(yd)";
      worksheet.getCell("K"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "K" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].ttl_marker)
       worksheet.getCell(y).numFmt = '0.00'; 
       worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
       }; 
       }

        //K-------------------------------------------------------------------------------
         worksheet.getCell("L"+a).value = "Plug1+2";
      worksheet.getCell("L"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("L"+b).value = "(inch)";
      worksheet.getCell("L"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
         
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "L" + ab
      
            var z = ab-a-2;
      worksheet.getCell(y).value = Number(this.rowexport[z].plug12)
      worksheet.getCell(y).numFmt = '0.00'; 
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
       }; 
       }

       //L------------------------------------------------------------------
            worksheet.getCell("M"+a).value = "%";
      worksheet.getCell("M"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("M"+b).value = 1.50/100;
       worksheet.getCell("M"+b).numFmt = '0.00%'; 
      worksheet.getCell("M"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

          this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "M" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].widthloss/100
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
       worksheet.getCell("M"+ty).value = "";
      worksheet.getCell("M"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


      //N----------------------------------------------------------
       worksheet.getCell("N"+a).value = "Length";
      worksheet.getCell("N"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("N"+b).value = "(yd)";
      worksheet.getCell("N"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "N" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].endless_length_yd)
       worksheet.getCell(y).numFmt = '0.00'; 
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
      this.check_value = ""
      if(this.multiple[x] == "Woven"){
       this.check_value =  this.end_loss_value[0].p1;
      }else if(this.multiple[x] == "Light-Middle Weight"){
       this.check_value =  this.end_loss_value[0].p2;
      }else if(this.multiple[x] == "fleece"){
        this.check_value =  this.end_loss_value[0].p3;
      }else{
        this.check_value = this.end_loss_value[0].p4;
      }

    worksheet.getCell("O"+b).value = this.check_value / 100;
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
       worksheet.getCell(y).value = this.rowexport[z].avg_end/100
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("P"+b).value = "(yd)";
      worksheet.getCell("P"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "P" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].spliceloss)
       worksheet.getCell(y).numFmt = '0.00'; 
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
      if(this.multiple[x] == "Woven"){
       worksheet.getCell("Q"+b).value = this.splice_value[0].p1 / 100;
      }else if(this.multiple[x] == "Light-Middle Weight"){
       worksheet.getCell("Q"+b).value = this.splice_value[0].p2 / 100;
      }else if(this.multiple[x] == "Fleece"){
        worksheet.getCell("Q"+b).value = this.splice_value[0].p3 / 100;
      }else{
        worksheet.getCell("Q"+b).value = this.splice_value[0].p4 / 100;
      }
 
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
       worksheet.getCell(y).value = this.rowexport[z].splicelossper/100
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("R"+b).value = "(yd)";
      worksheet.getCell("R"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
         this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "R" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].totalcutout)
       worksheet.getCell(y).numFmt = '0.00'; 
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
        if(this.multiple[x] == "Woven"){
       worksheet.getCell("S"+b).value = this.cut_out_value[0].p1 / 100;
      }else if(this.multiple[x] == "Light-Middle Weight"){
       worksheet.getCell("S"+b).value = this.cut_out_value[0].p2 / 100;
      }else if(this.multiple[x] == "Fleece"){
        worksheet.getCell("S"+b).value = this.cut_out_value[0].p3 / 100;
      }else{
        worksheet.getCell("S"+b).value = this.cut_out_value[0].p4 / 100;
      }
       worksheet.getCell("S"+b).numFmt = '0.00%'; 
      worksheet.getCell("S"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("T"+b).value = "(yd)";
      worksheet.getCell("T"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "T" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].rement)
       worksheet.getCell(y).numFmt = '0.00'; 
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       if(this.multiple[x] == "Woven"){
       worksheet.getCell("U"+b).value = this.remnants_loss_value[0].p1 / 100;
      }else if(this.multiple[x] == "Light-Middle Weight"){
       worksheet.getCell("U"+b).value = this.remnants_loss_value[0].p2 / 100;
      }else if(this.multiple[x] == "Fleece"){
        worksheet.getCell("U"+b).value = this.remnants_loss_value[0].p3 / 100;
      }else{
        worksheet.getCell("U"+b).value = this.remnants_loss_value[0].p4 / 100;
      }
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
       worksheet.getCell(y).value = this.rowexport[z].rement_per/100
       worksheet.getCell(y).numFmt = '0.00'; 
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("V"+b).value = "(yd)";
      worksheet.getCell("V"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "V" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = Number(this.rowexport[z].totallenthloss)
       worksheet.getCell(y).numFmt = '0.00'; 
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
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
      if(this.multiple[x] == "Woven"){
       worksheet.getCell("W"+b).value = this.total_woven / 100;
      }else if(this.multiple[x] == "Light-Middle Weight"){
       worksheet.getCell("W"+b).value = this.total_light / 100;
      }else if(this.multiple[x] == "Fleece"){
        worksheet.getCell("W"+b).value = this.total_fleece / 100;
      }else{
        worksheet.getCell("W"+b).value = this.total_special / 100;
      }
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
       worksheet.getCell(y).value = this.rowexport[z].totallenthlossper/100
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
      var d = this.length+2    
      var e = this.length+5
      var f = this.length+6
      var  g = this.length+7
      var  h = this.length+8
      var i = this.length+9
      var j = this.length+10
      var k = this.length+11
      var l = this.length+12 //endloss
      var m = this.length+13 //splice loss
      var n = this.length+14 //Cut - out Loss
      var o = this.length+15 //Remnants Loss
      var p = this.length+16 //Total Loss

      
      for(var xyz1 = 2; xyz1<p; xyz1++){
        worksheet.getRow(xyz1).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 11,
        bold: false
      };
    
      }

      let org = this.$q.localStorage.getItem("org");
     
      worksheet.getCell("A"+d).value = "NY"+org+ " : Weekly Total";
      worksheet.mergeCells("A"+d+":" +"I"+d);
      worksheet.getCell("A"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.result_sum_ttlmarker = 0
      for(var jd = 0; jd<this.keep_ttlmarker.length; jd++){
        this.result_sum_ttlmarker = Number(this.result_sum_ttlmarker)+Number(this.keep_ttlmarker[jd].result_ttlmarker)
      
      }
      
      worksheet.getCell("J"+d).value = "";
      worksheet.getCell("J"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
   

       /*  this.sum_plug12_inch = 0
        this.result_sum_plug12_inch = 0
       
      for(var kd = 0; kd<this.keep_ttlmarker.length; kd++){
         this.result_sum_plug12_inch = this.keep_ttlmarker[kd].result_ttlmarker * this.keep_plug12_inch[kd].result_plug12_inch
        

  
        this.sum_plug12_inch = Number(this.sum_plug12_inch) + Number(this.result_sum_plug12_inch)
      }
     
      
      console.log(this.sum_plug12_inch) */
       worksheet.getCell("K"+d).value =  Number(this.result_sum_ttlmarker)
      worksheet.getCell("K"+d).numFmt = '0.00';
       worksheet.getCell("K"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      /* this.result_sum_plug12_inch_l=0
      this.sum_plug12_inch_l =0
       for(var ld = 0; ld<this.keep_ttlmarker.length; ld++){
         this.result_sum_plug12_inch_l = this.keep_ttlmarker[ld].result_ttlmarker * this.keep_perplug12[ld].result_perplug12
        

  
        this.sum_plug12_inch_l = Number(this.sum_plug12_inch) + Number(this.result_sum_plug12_inch)
      }
      this.sum_plug12_inch_l = this.sum_plug12_inch_l/this.result_sum_ttlmarker
       */
      worksheet.getCell("L"+d).value = "";
      worksheet.getCell("L"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

  
          this.result_avg_end_plug12 = 0
          this.sum_avg_end_plug12 = 0
         for(var md = 0; md<this.keep_ttlmarker.length; md++){
          
         this.result_avg_end_plug12 = this.keep_ttlmarker[md].result_ttlmarker * this.keep_avg_loss_plug12.result_avg_end_plug12
        

  
        this.sum_avg_end_plug12 = Number(this.sum_avg_end_plug12) + Number(this.result_sum_plug12_inch)
        
      } 
     this.sum_avg_end_plug12 = this.sum_avg_end_plug12/this.result_sum_ttlmarker
      worksheet.getCell("M"+d).value = "";
      worksheet.getCell("M"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
   this.all_sum_length_yd = 0
   
     for(var azc = 0; azc<this.keep_avg_length_yd.length; azc++){
       this.all_sum_length_yd = Number(this.all_sum_length_yd)+ Number(this.keep_avg_length_yd[azc].all_result_avg_length)
      
      }   
      worksheet.getCell("N"+d).value = Number(this.all_sum_length_yd)
      worksheet.getCell("N"+d).numFmt = '0.00';
      worksheet.getCell("N"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
      this.sum_avg_per = this.all_sum_length_yd/this.result_sum_ttlmarker*100
  
       worksheet.getCell("O"+d).value = this.sum_avg_per/100
       worksheet.getCell("O"+d).numFmt = '0.00%';
       worksheet.getCell("O"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
      this.result_splice = 0
      for(var re_splice = 0; re_splice<this.keep_splice_Loss.length; re_splice++ ){
        this.result_splice =  Number(this.result_splice) + Number(this.keep_splice_Loss[re_splice].result_splice)
        
      }
     
       worksheet.getCell("P"+d).value = Number(this.result_splice)
       worksheet.getCell("P"+d).numFmt = '0.00';
       worksheet.getCell("P"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.sum_result_splice_per = this.result_splice/this.result_sum_ttlmarker*100
      
       worksheet.getCell("Q"+d).value = this.sum_result_splice_per/100
       worksheet.getCell("Q"+d).numFmt = '0.00%';
       worksheet.getCell("Q"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      

       this.result_sum_cut_loss = 0
      for(var re_cut_out = 0; re_cut_out<this.keep_cut_out_loss.length; re_cut_out++ ){
        this.result_sum_cut_loss =  Number(this.result_sum_cut_loss) + Number(this.keep_cut_out_loss[re_cut_out].result_cut_out_loss)
        
      }
     

    
       worksheet.getCell("R"+d).value =  Number(this.result_sum_cut_loss)
       worksheet.getCell("R"+d).numFmt = '0.00';
       worksheet.getCell("R"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
     
      this.sum_cut_loss_out_per = this.result_sum_cut_loss/this.result_sum_ttlmarker*100
     
       worksheet.getCell("S"+d).value =  this.sum_cut_loss_out_per/100
       worksheet.getCell("S"+d).numFmt = '0.00%';
       worksheet.getCell("S"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


       this.result_sum_remnants = 0
      for(var re_rem = 0; re_rem<this.keep_remnants.length; re_rem++ ){
        this.result_sum_remnants =  Number(this.result_sum_remnants) + Number(this.keep_remnants[re_rem].result_remnants)
       
        
      }
       worksheet.getCell("T"+d).value = (this.result_sum_remnants)
       worksheet.getCell("T"+d).numFmt = '0.00';
       worksheet.getCell("T"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.sum_remnants_per = this.result_sum_remnants/this.result_sum_ttlmarker*100  
       worksheet.getCell("U"+d).value = this.sum_remnants_per/100
       worksheet.getCell("U"+d).numFmt = '0.00%';
       worksheet.getCell("U"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.sum_total_length = 0
      for(var total_sum = 0; total_sum<this.keep_total_loss.length; total_sum++ ){
        this.sum_total_length =  Number(this.sum_total_length) + Number(this.keep_total_loss[total_sum].result_total_length)
      }

       worksheet.getCell("V"+d).value = Number(this.sum_total_length)
       worksheet.getCell("V"+d).numFmt = '0.00';
       worksheet.getCell("V"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.sum_total_length_per = this.sum_total_length/this.result_sum_ttlmarker*100
       worksheet.getCell("W"+d).value = this.sum_total_length_per/100
       worksheet.getCell("W"+d).numFmt = '0.00%';
       worksheet.getCell("W"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
      worksheet.getCell("Y3").value = "AVG.Marker Length By Fabric Type";
      worksheet.mergeCells("Y3:AB3");
      worksheet.getCell("Y3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Y3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("AC3").value = "All Fabric Type";
      worksheet.mergeCells("AC3:AC4");
      worksheet.getCell("AC3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("AC3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
  
       worksheet.getCell("Y4").value = "T1";
       worksheet.getCell("Y4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Y4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("Y5").value = "";
      worksheet.getCell("Y5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("Z4").value = "T2";
         worksheet.getCell("Z4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Z4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("Z5").value = "";
      worksheet.getCell("Z5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("AA4").value = "T3";
        worksheet.getCell("AA4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("AA4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

           worksheet.getCell("AA5").value = ""
      worksheet.getCell("AA5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

           worksheet.getCell("AB4").value = "T4";
         worksheet.getCell("AB4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("AB4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

    
         worksheet.getCell("AB5").value = "";
      worksheet.getCell("AB5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      

      console.log(this.keep_avg_marker_length)
      for(var ax = 0; ax<this.keep_avg_marker_length.length; ax++){
      
      if(this.keep_avg_marker_length[ax].org == "1"){
    
      worksheet.getCell("Y5").value = Number(this.keep_avg_marker_length[ax].value);
       worksheet.getCell("Y5").numFmt ='0.00';
       worksheet.getCell("Y5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Y5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }
       
      
      }
   
       for(var ax = 0; ax<this.keep_avg_marker_length.length; ax++){
     
      if(this.keep_avg_marker_length[ax].org == "2"){
    
         worksheet.getCell("Z5").value = Number(this.keep_avg_marker_length[ax].value);
         worksheet.getCell("Z5").numFmt = '0.00';
         worksheet.getCell("Z5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Z5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }
      
      
       }
        for(var ax = 0; ax<this.keep_avg_marker_length.length; ax++){
      if(this.keep_avg_marker_length[ax].org == "3"){

         worksheet.getCell("AA5").value = Number(this.keep_avg_marker_length[ax].value)
         worksheet.getCell("AA5").numFmt = '0.00';
         worksheet.getCell("AA5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("AA5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }
      
      
        }

        for(var ax = 0; ax<this.keep_avg_marker_length.length; ax++){
      if(this.keep_avg_marker_length[ax].org == "4"){

         worksheet.getCell("AB5").value =Number(this.keep_avg_marker_length[ax].value);
         worksheet.getCell("AB5").numFmt = '0.00';
         worksheet.getCell("AB5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("AB5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }
        }
   




     this.sum_keep_sum_x = 0
      for(var ax = 0; ax<this.keep_sum_x.length; ax++){
          this.sum_keep_sum_x = Number(this.sum_keep_sum_x) + Number(this.keep_sum_x[ax].value)
      }
      this.sum_keep_ttl_sum_x = 0
      for(var ay = 0; ay<this.keep_sum_ttl_x.length; ay++){
          this.sum_keep_ttl_sum_x = Number(this.sum_keep_ttl_sum_x) + Number(this.keep_sum_ttl_x[ay].value)
      }

      this.result_sum_x_ttl = 0
      this.result_sum_x_ttl = this.sum_keep_sum_x / this.sum_keep_ttl_sum_x
    
       worksheet.getCell("AC5").value = Number(this.result_sum_x_ttl)
       worksheet.getCell("AC5").numFmt = '0.00';
        worksheet.getCell("AC5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("AC5").border = {
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
        saveAs(blob, "report_weekly.xlsx");
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