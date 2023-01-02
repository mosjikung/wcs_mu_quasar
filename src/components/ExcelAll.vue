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
            
            class="q-px-xl"
            v-model="org"
            :options="option_org"
            style="width:150px"
            disable
            
            label="Chose Org"
          />
    
             <q-btn
      size="md"
      dense
      class="q-px-xl q-py-xs "
      color="primary"
      label="Export"
      @click="Export()"
    />
  </div>
</template>

<script>
import {ref} from 'vue'
import {date} from 'quasar'
import axios from 'axios'
import Detail from '../components/Detail.vue'
//import ExcelDaily from '../components/Excel_Daily.vue'
import { useQuasar, QSpinnerGears } from 'quasar'
import { onBeforeUnmount } from 'vue'
import * as Excel from "exceljs";
import { saveAs } from "file-saver";
export default {
  data() {
    return {
      start:"",
      end:"",
      org:"",
      rowexport:[],
      column_row:[
        {
          colname:"A",
        },
        {
          colname:"B",
        },
        {
          colname:"C",
        },
        {
          colname:"D",
        },
        {
          colname:"E",
        },
        {
          colname:"F",
        },
        {
          colname:"G",
        },
        {
          colname:"H",
        },
        {
          colname:"I",
        },
        {
          colname:"J",
        },
        {
          colname:"K",
        },
        {
          colname:"L",
        },
        {
          colname:"M",
        },
        {
          colname:"N",
        },
        {
          colname:"O",
        },
        {
          colname:"P",
        },
        {
          colname:"Q",
        },
        {
          colname:"R",
        },
        {
          colname:"S",
        },
        {
          colname:"T",
        },
        {
          colname:"U",
        },
        {
          colname:"V",
        },
        {
          colname:"W",
        },
        {
          colname:"X",
        },
        {
          colname:"Y",
        },
        {
          colname:"Z",
        },
        {
          colname:"AA",
        },
        {
          colname:"AB",
        },
        {
          colname:"AC",
        },
        {
          colname:"AD",
        },
        {
          colname:"AE",
        },
        {
          colname:"AF",
        },
        {
          colname:"AG",
        },
        {
          colname:"AH",
        },
        {
          colname:"AI",
        },
        {
          colname:"AJ",
        },
        {
          colname:"AK",
        },

        
      ],
      end_loss_value:[],
      splice_value:[],
      cut_out_value:[],
      remnant_value:[],




    };
  },
   mounted() {
   
      this.end_loss_value=[],
      this.splice_value=[],
      this.cut_out_value=[],
      this.remnant_value=[],
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
  methods: {
       Export() {

        const params = new FormData()

        params.append("start",this.start_date)
        params.append("end",this.end_date)
        let org = this.$q.localStorage.getItem("org");
        params.append("org",this.org)
       

        axios({
        method:'post',
        url: this.$api_url + '/mainso.php/export_data',
        data:params,
      }).then((resp)=>{
        

             if (resp.data.data.length > 0) {
               function check_null(val){
                if(val == '' &&  val == null && val == "NaN"){
                  return 0
                }else{
                  return val
                }
               }
              resp.data.data.forEach((e) => {
                this.rowexport.push({
                  mu_date: e.MU_DATE,
                  gm_table:e.GM_TABLE,
                  so_no_doc:e.SO_NO_DOC,
                  customer:e.CUSTOMER_NAME,
                  gm_no:e.GM_NO,
                  fabric_type:e.FABRIC_TYPE,
                  issue: e.ISSUE,
                  issue_roll :e.ROLL_NO,
                  avg_width:e.AVG_WIDTH,
                  gmm2: e.WEIGHT_M2,
                  gmyd: e.WEIGHT_YD,
                  width_inch : e.WIDTH_INCH,
                  length_ydb:e.YARD,
                  length_inch: e.LENGTH_INCH,
                  obs_width: e.OBS_WIDTH,
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
                  remnants: e.REMNANTS_WEIGHT,
                  widthplug1: e.PLUG1_GRAM,
                  widthplug2: e.PLUG2_GRAM,
                  avgglueside: e.GLUE_SIDE,
                  ttl_marker: parseFloat(e.LENGTH_YD*e.PLY).toFixed(2),
                  yard_roll:e.YARD_ROLL,
                  length_yd:e.LENGTH_YD,
                  obswidth:e.OBS_WIDTH,
                  ttlobs_widthloss:e.OBS_WIDTH,
                  end1_2:check_null(e.END1_2),
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
                  spliceloss:e.SPLICE_LOSS_YD,
                  splicelossper:parseFloat(e.SPLICE_LOSS_PER).toFixed(2),
                  overlap:e.OVER_LAB_YD,
                  overlapper:parseFloat(e.OVER_LAB_PER).toFixed(2),
                  defect:parseFloat(e.DEFECT_YD).toFixed(2),
                  defectper:parseFloat(e.DEFECT_PER).toFixed(2),
                  totalcutout:parseFloat(e.TOTAL_CUTOUT_YD).toFixed(2),
                  totalcutoutper:parseFloat(e.TOTAL_CUTOUT_PER).toFixed(2),
                  rement:parseFloat(e.REMENT_LOSS_YD).toFixed(2),
                  rement_per:parseFloat(e.REMENT_LOSS_PER).toFixed(2),
                  totallenthloss:parseFloat(e.TOTAL_LENGTH_LOSS_YD).toFixed(2),
                  totallenthlossper:parseFloat(e.TOTAL_LEN_LOSS_PER).toFixed(2),
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

                });
               
              });
              const workbook = new Excel.Workbook();
      workbook.creator = "Nanyang";
      workbook.lastModifiedBy = "Admin";
      workbook.created = new Date(2021, 8, 30);
      workbook.modified = new Date();
      workbook.lastPrinted = new Date(2021, 7, 27);
      const worksheet = workbook.addWorksheet("New Sheet");
      //คอลัม

       worksheet.columns = [ {key: 'A', width: 10},{key: 'B', width: 7.63},{key: 'C', width: 13},{key: 'D', width: 40},{key: 'E', width: 13}
       ,{key: 'F', width: 36.4},{key: 'G', width: 8.7},{key: 'H', width: 6.7},{key: 'I', width: 6.7},{key: 'J', width: 12},{key: 'K', width: 12},
       {key: 'L', width: 6.7},{key: 'M', width: 45},{key: 'N', width: 6.7},{key: 'O', width: 6.7},{key: 'P', width: 6.7},{key: 'Q', width: 6.7},
        {key: 'R', width: 6.7},{key: 'S', width: 12},{key: 'T', width: 6.7},{key: 'U', width: 6.7},{key: 'V', width: 6.7},{key: 'W', width: 6.7},{key: 'X', width: 6.7}]; 

       worksheet.getCell("A1").value = "SPREAD LOSS DATA :  NY"+ this.org;
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

      worksheet.getCell("A3").value = "date";
      worksheet.getCell("A3").alignment = { horizontal: "center" };
      worksheet.getCell("A3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("Y5").value = "code";
      worksheet.getCell("Y5").alignment = { horizontal: "center" };
      worksheet.getCell("Y5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("Z5").value = "casue";
      worksheet.getCell("Z5").alignment = { horizontal: "center" };
      worksheet.getCell("Z5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("AA5").value = "code";
      worksheet.getCell("AA5").alignment = { horizontal: "center" };
      worksheet.getCell("AA5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("AB5").value = "casue";
      worksheet.getCell("AB5").alignment = { horizontal: "center" };
      worksheet.getCell("AB5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

         worksheet.getCell("AC5").value = "code";
      worksheet.getCell("AC5").alignment = { horizontal: "center" };
      worksheet.getCell("AC5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("AD5").value = "casue";
      worksheet.getCell("AD5").alignment = { horizontal: "center" };
      worksheet.getCell("AD5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

              worksheet.getCell("AE5").value = "code";
      worksheet.getCell("AE5").alignment = { horizontal: "center" };
      worksheet.getCell("AE5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("AF5").value = "casue";
      worksheet.getCell("AF5").alignment = { horizontal: "center" };
      worksheet.getCell("AF5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

              worksheet.getCell("AG5").value = "code";
      worksheet.getCell("AG5").alignment = { horizontal: "center" };
      worksheet.getCell("AG5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("AH5").value = "casue";
      worksheet.getCell("AH5").alignment = { horizontal: "center" };
      worksheet.getCell("AH5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      //merge
      worksheet.mergeCells("A1", "AK2");
      worksheet.getCell("A1").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.mergeCells("B3:E3");
      worksheet.getCell("B3").value = this.rowexport[0].mu_date;
      worksheet.getCell("B3").alignment = { horizontal: "center" };
      worksheet.getCell("B3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.mergeCells("F3:G3");
      worksheet.getCell("F3").value = "Fabric Type";
      worksheet.getCell("F3").alignment = { horizontal: "center" };
      worksheet.getCell("F3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.mergeCells("H3:J3");
      worksheet.getCell("H3").value = "Marker";
      worksheet.getCell("H3").alignment = { horizontal: "center" };
      worksheet.getCell("H3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.mergeCells("K3:L3");
      worksheet.getCell("K3").value = "Width Loss";
      worksheet.getCell("K3").alignment = { horizontal: "center" };
      worksheet.getCell("K3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.mergeCells("M3:M5");
      worksheet.getCell("M3").key ="Width Loss Cause"
      worksheet.getCell("M3").value = "Width Loss Cause";
      worksheet.getCell("M3").alignment = {
        horizontal: "center",
        vertical: "middle",
      };
      worksheet.getCell("M3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      worksheet.getCell.size="100"

      worksheet.mergeCells("N3:P3");
      worksheet.getCell("N3").value = "AVG. End Loss";
      worksheet.getCell("N3").alignment = { horizontal: "center" };
      worksheet.getCell("N3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.mergeCells("Q3:R3");
      worksheet.getCell("Q3").value = "Splice Loss";
      worksheet.getCell("Q3").alignment = { horizontal: "center" };
      worksheet.getCell("Q3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.mergeCells("S3:T3");
      worksheet.getCell("S3").value = "Cut out Loss";
      worksheet.getCell("S3").alignment = { horizontal: "center" };
      worksheet.getCell("S3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.mergeCells("U3:V3");
      worksheet.getCell("U3").value = "Remenants Loss";
      worksheet.getCell("U3").alignment = { horizontal: "center" };
      worksheet.getCell("U3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.mergeCells("W3:X3");
      worksheet.getCell("W3").value = "Total Length Loss";
      worksheet.getCell("W3").alignment = { horizontal: "center" };
      worksheet.getCell("W3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       worksheet.mergeCells("Y3:Z4");
      worksheet.getCell("Y3").value = "Width Loss";
      worksheet.getCell("Y3").alignment = { horizontal: "center" };
      worksheet.getCell("Y3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.mergeCells("AA3:AB4");
      worksheet.getCell("AA3").value = "End Loss";
      worksheet.getCell("AA3").alignment = { horizontal: "center" };
      worksheet.getCell("AA3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

         worksheet.mergeCells("AC3:AD4");
      worksheet.getCell("AC3").value = "Splice Loss";
      worksheet.getCell("AC3").alignment = { horizontal: "center" };
      worksheet.getCell("AC3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

          worksheet.mergeCells("AE3:AF4");
      worksheet.getCell("AE3").value = "Cut Out Loss";
      worksheet.getCell("AE3").alignment = { horizontal: "center" };
      worksheet.getCell("AE3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

          worksheet.mergeCells("AG3:AH4");
      worksheet.getCell("AG3").value = "Remnants Loss";
      worksheet.getCell("AG3").alignment = { horizontal: "center" };
      worksheet.getCell("AG3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.mergeCells("AI3:AI5");
      worksheet.getCell("AI3").value = "Ply";
      worksheet.getCell("AI3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("AI3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

          worksheet.mergeCells("AJ3:AJ5");
      worksheet.getCell("AJ3").value = "Roll";
      worksheet.getCell("AJ3").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("AJ3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("AK3").value = "Balance";
      worksheet.getCell("AK3").alignment = { horizontal: "center" };
      worksheet.getCell("AK3").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

          worksheet.mergeCells("AK4:AK5");
      worksheet.getCell("AK4").value = "Issue-Use(KG)";
      worksheet.getCell("AK4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("AK4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("A4").value = "Date";
      worksheet.mergeCells("A4:A5");
      worksheet.getCell("A4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

         worksheet.getCell("B4").value = "FL";
      worksheet.mergeCells("B4:B5");
      worksheet.getCell("B4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("B4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("C4").value = "S/O no.";
      worksheet.mergeCells("C4:C5");
      worksheet.getCell("C4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("C4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("D4").value = "Customer";
      worksheet.mergeCells("D4:D5");
      worksheet.getCell("D4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("D4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       worksheet.getCell("E4").value = "GM08 no.";
      worksheet.mergeCells("E4:E5");
      worksheet.getCell("E4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("E4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       worksheet.getCell("F4").value = "Woven";
      worksheet.mergeCells("F4:F5");
      worksheet.getCell("F4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("F4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("G4").value = "Observe";
       worksheet.getCell("G4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("G4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("G5").value = "Width";
       worksheet.getCell("G5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("G5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("H4").value = "Width";
       worksheet.getCell("H4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("H5").value = "Inch";
      worksheet.getCell("H5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("I4").value = "Lenght";
       worksheet.getCell("I4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("I5").value = "(YD)";
       worksheet.getCell("I5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("J4").value = "Ttl Marker";
       worksheet.getCell("J4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("J5").value = "Lenght(YD)";
       worksheet.getCell("J5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("K4").value = "Plug 1+2";
       worksheet.getCell("K4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("K5").value = "(inch)";
       worksheet.getCell("K5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("L4").value = "%";
       worksheet.getCell("L4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("L5").value = "1.50%";
       worksheet.getCell("L5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("N4").value = "Plug1+2";
       worksheet.getCell("N4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("N5").value = "(inch)";
       worksheet.getCell("N5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("O4").value = "Lenght";
       worksheet.getCell("O4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("O5").value = "(yd)";
       worksheet.getCell("O5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("P4").value = "%";
       worksheet.getCell("P4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

    worksheet.getCell("P5").value = "0.15%";
       worksheet.getCell("P5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

  worksheet.getCell("Q4").value = "Lenght";
       worksheet.getCell("Q4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("Q5").value = "(yd)";
       worksheet.getCell("Q5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("R4").value = "%";
       worksheet.getCell("R4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

    worksheet.getCell("R5").value = "0.20%";
       worksheet.getCell("R5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("S4").value = "Lenght";
       worksheet.getCell("S4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("S5").value = "(yd)";
       worksheet.getCell("S5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("T4").value = "%";
       worksheet.getCell("T4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

    worksheet.getCell("T5").value = "0.00%";
       worksheet.getCell("T5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("U4").value = "Lenght";
       worksheet.getCell("U4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("U5").value = "(yd)";
       worksheet.getCell("U5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("V4").value = "%";
       worksheet.getCell("V4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

    worksheet.getCell("V5").value = "0.08%";
       worksheet.getCell("V5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("W4").value = "Lenght";
       worksheet.getCell("W4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("W5").value = "(yd)";
       worksheet.getCell("W5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("X4").value = "%";
       worksheet.getCell("X4").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("X4").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

    worksheet.getCell("X5").value = "0.43%";
       worksheet.getCell("X5").alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("X5").border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       this.length =  resp.data.data.length + 6
      
       for(var x=6; x<this.length; x++){
      var y  = "A" + x
      
            var z = x-6;
       worksheet.getCell(y).value = this.rowexport[z].mu_date;
      
      

            }

       for(var x=6; x<this.length; x++){
      var y  = "B" + x
     
            var z = x-6;
       worksheet.getCell(y).value = this.rowexport[z].gm_table;
      
      

            }

             for(var x=6; x<this.length; x++){
      var y  = "C" + x
      
            var z = x-6;
       worksheet.getCell(y).value = this.rowexport[z].gm_no;
      
      

            }
        for(var x=6; x<this.length; x++){
      var y  = "D" + x
      
            var z = x-6;
       worksheet.getCell(y).value = this.rowexport[z].customer;
      
      

            }

          for(var x=6; x<this.length; x++){
      var y  = "E" + x
      
            var z = x-6;
       worksheet.getCell(y).value = this.rowexport[z].gm_no;
      
      

            }

        for(var x=6; x<this.length; x++){
      var y  = "F" + x
      
            var z = x-6;
       worksheet.getCell(y).value = this.rowexport[z].fabric_type;
      
      

            }

      for(var x=6; x<this.length; x++){
      var y  = "G" + x
      
            var z = x-6;
       worksheet.getCell(y).value = Number(this.rowexport[z].obs_width);
       worksheet.getCell(y).numFmt = "0.00";
      
      

            }

       for(var x=6; x<this.length; x++){
      var y  = "H" + x
      
            var z = x-6;
       worksheet.getCell(y).value = Number(this.rowexport[z].width_inch);
       worksheet.getCell(y).numFmt = "0.00";
      
      

            }

       for(var x=6; x<this.length; x++){
      var y  = "I" + x
      
            var z = x-6;
       worksheet.getCell(y).value = Number(this.rowexport[z].length_yd);
       worksheet.getCell(y).numFmt = "0.00";
      

            }

        for(var x=6; x<this.length; x++){
      var y  = "J" + x
      
            var z = x-6;
       worksheet.getCell(y).value = Number(this.rowexport[z].ttl_marker);
       worksheet.getCell(y).numFmt = "0.00";
      
      

            }

        for(var x=6; x<this.length; x++){
      var y  = "K" + x
      
            var z = x-6;
       worksheet.getCell(y).value = Number(this.rowexport[z].plug12);
       worksheet.getCell(y).numFmt = "0.00";
      
      

            }

       for(var x=6; x<this.length; x++){
      var y  = "L" + x
      
            var z = x-6;
       worksheet.getCell(y).value = this.rowexport[z].widthloss/100;
       worksheet.getCell(y).numFmt = "0.00%";
      
      

            }

       for(var x=6; x<this.length; x++){
      var y  = "M" + x
      
            var z = x-6;
       worksheet.getCell(y).value = this.rowexport[z].cause;
      
      

            }

       for(var x=6; x<this.length; x++){
      var y  = "N" + x
      
            var z = x-6;
       worksheet.getCell(y).value = Number(this.rowexport[z].plug12);
       worksheet.getCell(y).numFmt = "0.00";
      
      

            }

     

       for(var x=6; x<this.length; x++){
      var y  = "O" + x
      
            var z = x-6;
       worksheet.getCell(y).value = Number(this.rowexport[z].length_yd);
       worksheet.getCell(y).numFmt = "0.00";
      

            }

     

      for(var x=6; x<this.length; x++){
      var y  = "P" + x
      
            var z = x-6;
       worksheet.getCell(y).value = this.rowexport[z].avg_end/100;
       worksheet.getCell(y).numFmt = "0.00%";
      
      

            }

      for(var x=6; x<this.length; x++){
      var y  = "Q" + x
      
            var z = x-6;
       worksheet.getCell(y).value = Number(this.rowexport[z].sectionlosslenth);
       worksheet.getCell(y).numFmt = "0.00";
      
      

            }

       for(var x=6; x<this.length; x++){
      var y  = "R" + x
      
            var z = x-6;
       worksheet.getCell(y).value = this.rowexport[z].sectionlossper/100;
       worksheet.getCell(y).numFmt = "0.00%";
      
      

            }

       for(var x=6; x<this.length; x++){
      var y  = "S" + x
      
            var z = x-6;
             if(this.rowexport[z].totalcutout == "NaN"){
              this.rowexport[z].totalcutout = "";
            }
       worksheet.getCell(y).value = Number(this.rowexport[z].totalcutout);
       worksheet.getCell(y).numFmt = "0.00";
      
      

            }

       for(var x=6; x<this.length; x++){
      var y  = "T" + x
      
            var z = x-6;
             if(this.rowexport[z].totalcutoutper == "NaN"){
              this.rowexport[z].totalcutoutper = "";
            }
       worksheet.getCell(y).value = this.rowexport[z].totalcutoutper/100;
      worksheet.getCell(y).numFmt = "0.00%";
      

            }

       for(var x=6; x<this.length; x++){
      var y  = "U" + x
      
            var z = x-6;

            if(this.rowexport[z].rement == "NaN"){
              this.rowexport[z].rement = "";
            }
       worksheet.getCell(y).value = Number(this.rowexport[z].rement);
       worksheet.getCell(y).numFmt = "0.00";
      
      

            }

       for(var x=6; x<this.length; x++){
      var y  = "V" + x
      
            var z = x-6;
             if(this.rowexport[z].rement_per == "NaN"){
              this.rowexport[z].rement_per = "";
            }
       worksheet.getCell(y).value = this.rowexport[z].rement_per/100;
       worksheet.getCell(y).numFmt = "0.00%";
      
      

            }

       for(var x=6; x<this.length; x++){
      var y  = "W" + x
      
            var z = x-6;
            if(this.rowexport[z].totallenthloss == "NaN"){
              this.rowexport[z].totallenthloss = "";
            }
       worksheet.getCell(y).value = Number(this.rowexport[z].totallenthloss);
       worksheet.getCell(y).numFmt = "0.00";
      
      
       }

         for(var x=6; x<this.length; x++){
      var y  = "X" + x
      
            var z = x-6;
             if(this.rowexport[z].totallenthlossper == "NaN"){
              this.rowexport[z].totallenthlossper = "";
            }
       worksheet.getCell(y).value = this.rowexport[z].totallenthlossper/100;
       worksheet.getCell(y).numFmt = "0.00%";
      
      
      } 

       for(var x=6; x<this.length; x++){
      var y  = "Y" + x
      
            var z = x-6;
             if(this.rowexport[z].widthloss_code == "NaN"){
              this.rowexport[z].widthloss_code = "";
            }
       worksheet.getCell(y).value = this.rowexport[z].widthloss_code;
      
      
      }

        for(var x=6; x<this.length; x++){
      var y  = "Z" + x
      
            var z = x-6;
             if(this.rowexport[z].widthloss_cause == "NaN"){
              this.rowexport[z].widthloss_cause = "";
            }
       worksheet.getCell(y).value = this.rowexport[z].widthloss_cause;
      
      
      }

       for(var x=6; x<this.length; x++){
      var y  = "AA" + x
      
            var z = x-6;
             if(this.rowexport[z].avg_end_code == "NaN"){
              this.rowexport[z].avg_end_code = "";
            }
       worksheet.getCell(y).value = this.rowexport[z].avg_end_code;
      
       
      }

        for(var x=6; x<this.length; x++){
      var y  = "AB" + x
      
            var z = x-6;
             if(this.rowexport[z].avg_end_cause == "NaN"){
              this.rowexport[z].avg_end_cause = "";
            }
       worksheet.getCell(y).value = this.rowexport[z].avg_end_cause;
      
      
      }

        for(var x=6; x<this.length; x++){
      var y  = "AC" + x
      
            var z = x-6;
             if(this.rowexport[z].per_splice_loss_code == "NaN"){
              this.rowexport[z].per_splice_loss_code = "";
            }
       worksheet.getCell(y).value = this.rowexport[z].per_splice_loss_code;
      
      
      }

              for(var x=6; x<this.length; x++){
      var y  = "AD" + x
      
            var z = x-6;
             if(this.rowexport[z].per_splice_loss_cause == "NaN"){
              this.rowexport[z].per_splice_loss_cause = "";
            }
       worksheet.getCell(y).value = this.rowexport[z].per_splice_loss_cause;
      
      
      }

              for(var x=6; x<this.length; x++){
      var y  = "AE" + x
      
            var z = x-6;
             if(this.rowexport[z].total_cut_out_per_code == "NaN"){
              this.rowexport[z].total_cut_out_per_code = "";
            }
       worksheet.getCell(y).value = this.rowexport[z].total_cut_out_per_code;
      
      
      }

              for(var x=6; x<this.length; x++){
      var y  = "AF" + x
      
            var z = x-6;
             if(this.rowexport[z].total_cut_out_per_cause == "NaN"){
              this.rowexport[z].total_cut_out_per_cause = "";
            }
       worksheet.getCell(y).value = this.rowexport[z].total_cut_out_per_cause;
      
      
      }

              for(var x=6; x<this.length; x++){
      var y  = "AG" + x
      
            var z = x-6;
             if(this.rowexport[z].remnants_loss_code == "NaN"){
              this.rowexport[z].remnants_loss_code = "";
            }
       worksheet.getCell(y).value = this.rowexport[z].remnants_loss_code;
      
      
      }

              for(var x=6; x<this.length; x++){
      var y  = "AH" + x
      
            var z = x-6;
             if(this.rowexport[z].remnants_loss_cause == "NaN"){
              this.rowexport[z].remnants_loss_cause = "";
            }
       worksheet.getCell(y).value = this.rowexport[z].remnants_loss_cause;
      
      
      }

              for(var x=6; x<this.length; x++){
      var y  = "AI" + x
      
            var z = x-6;
             if(this.rowexport[z].ply == "NaN"){
              this.rowexport[z].ply = "";
            }
       worksheet.getCell(y).value = Number(this.rowexport[z].ply);
       worksheet.getCell(y).numFmt = "0.00";
      
     
      }

               for(var x=6; x<this.length; x++){
      var y  = "AJ" + x
      
            var z = x-6;
             if(this.rowexport[z].roll == "NaN"){
              this.rowexport[z].roll = "";
            }
       worksheet.getCell(y).value = Number(this.rowexport[z].roll);
       worksheet.getCell(y).numFmt = "0.00";
      
       
      }
               for(var x=6; x<this.length; x++){
      var y  = "AK" + x
      
            var z = x-6;
             if(this.rowexport[z].balance == "NaN"){
              this.rowexport[z].balance = "";
            }
       worksheet.getCell(y).value = Number(this.rowexport[z].balance);
       worksheet.getCell(y).numFmt = "0.00";
      
      
      }
      
       for(var ax = 0; ax<this.column_row.length; ax++){
        for(var bz = 3; bz<this.rowexport.length+6; bz++){
        worksheet.getCell(this.column_row[ax].colname+[bz]).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
        worksheet.getCell(this.column_row[ax].colname+[bz]).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        worksheet.getCell(this.column_row[ax].colname+[bz]).alignment = {
          horizontal: "center",
          vertical: "middle",
        };

        }
      } 

      workbook.xlsx.writeBuffer().then((data) => {
        const blob = new Blob([data], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8",
        });
        saveAs(blob, "Spread Loss Data NY"+this.org+".xlsx");
      }); 

            }  

      }).catch((error)=>{
        console.log(error)
      }) 

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
};
</script>

