<template>
<div class="bg-blue-10 text-white">
      <q-toolbar>
        <q-btn flat round dense />

        <q-toolbar-title>Weekly Wla1-Data</q-toolbar-title>

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
  
    <q-select
      class="q-pa-xl"
      v-model="org"
      :options="option_org"
      style="width: 200px"
      disable
      label="Chose Org"
    />

    <div class="myclass">  
      <q-btn
      size="lg"
      dense
      color="primary"
      label="Export"
      @click="exportexcel()"
    /></div>
      </div>
    <br>
    <div class="row justify-center">
    <q-input class="q-pa-xl" filled v-model="start_date_for_graph" label="Start_date">
      <template v-slot:append>
        <q-icon name="event" class="cursor-pointer">
          <q-popup-proxy
            ref="qDateProxy"
            transition-show="scale"
            transition-hide="scale"
          >
            <q-date v-model="start_date_for_graph" mask="DD/MM/YYYY">
              <div class="row items-center justify-end">
                <q-btn v-close-popup label="Close" color="primary" flat />
              </div>
            </q-date>
          </q-popup-proxy>
        </q-icon>
      </template>
    </q-input>
  
    <q-input class="q-pa-xl" filled v-model="end_date_for_graph" label="End Date">
      <template v-slot:append>
        <q-icon name="event" class="cursor-pointer">
          <q-popup-proxy
            ref="qDateProxy"
            transition-show="scale"
            transition-hide="scale"
          >
            <q-date v-model="end_date_for_graph" mask="DD/MM/YYYY">
              <div class="row items-center justify-end">
                <q-btn v-close-popup label="Close" color="primary" flat />
              </div>
            </q-date>
          </q-popup-proxy>
        </q-icon>
      </template>
    </q-input>
    <p style="color:red" class="q-px-md q-pt-xl">***Chose Date For Graph</p>
      </div>
    <br>

      <div class="q-pa-xl" style="min-width:100%;" >
  <q-banner inline-actions rounded class="bg-orange-4 text-white" >
    <h5 class="q-mt-xs" >Graph For WLA1-Data</h5>
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

    

</template>

<script>
import { ref } from "vue";
import { date } from "quasar";
import axios from "axios";
import Detail from "../components/Detail.vue";
import { useQuasar, QSpinnerGears } from "quasar";
import { onBeforeUnmount } from "vue";
import * as Excel from "exceljs";
import { saveAs } from "file-saver";
import moment from "moment";
import * as echarts from 'echarts';
export default {
  data() {
    return {
      multiple: ref(null),
      org: "",
      options: ["Woven", "Light-Middle Weight", "Fleece", "Special"],
      rows: 6,
      start: "",
      end: "",
      start_date: "",
      end_date: "",
      posts:[],
      date_use_data:[],
      date_use_data2:[],
      date_start: "2021/10/25",
      data_count: 10,
      all_date: [],
      rowexport: [],
      rowexport2: [],
      data_start_date:[],
      all_data_arr:[],
      keep_target_widthloss:[],
      total_keep_target_graph:[],
      all_ttl_marker_arr:[],
      all_data_arr2:[],
      start_date_for_graph:"",
      end_date_for_graph:"",
      date_data_for_graph:[],
      total_keep_target_graph2:[],
      date_data_for_graph2:[],
      keep_target:[],
      keep_target2:[],
      keep_target_for_graph:[],
      keep_target_for_graph2:[],
      keep_acc_act_widthloss:[],
      keep_acc_act_widthloss_graph:[],
      acc_act_widthlossx:[],
      date_use_data_graph:[],
      date_use_data_graph2:[],
      x: 0,
      target:1.50
    
    };
  },
  mounted() {
    let org = this.$q.localStorage.getItem("org");
    this.org = org;

       axios.get(this.$api_url + "/mainso.php/get_date2")
      .then((resp)=>{
        this.start = (resp.data.data[0].START_DATE)
      }).catch((error)=>{
        console.log(error)
      })

     /*    axios
    .get(this.$api_url + "/somain.php/detail_start_date_first")
    .then((resp)=>{
        console.log(resp.data)
        this.posts = resp.data.data[0]
        this.start_date = resp.data.data[0].START_DATE
    }) */
   
  },
  methods: {
    checkdata() {
      for (var x = 0; x < this.multiple.length; x++) {
       
      }
    },
    async exportexcel() {
       this.date_use_data_graph = []
       this.total_keep_target_graph = []
       this.total_keep_target_graph2 = []
       this.date_data_for_graph = []
       this.keep_target=[]
       this.keep_target_for_graph=[]
       this.keep_acc_act_widthloss=[]
       this.keep_acc_act_widthloss_graph=[]
       this.acc_act_widthlossx=[]
        this.$q.loading.show({
          spinner: QSpinnerGears,
          spinnerColor: 'wthite',
          spinnerSize: 180,
          backgroundColor: 'black',
         
        })
      this.all_ttl_marker_arr = []
      this.all_data_arr = []
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
        { key: "T", width: 6.7 },
        { key: "U", width: 6.7 },
        { key: "V", width: 6.7 },
        { key: "W", width: 6.7 },
        { key: "X", width: 6.7 },
        { key: "AC", width: 15.13 },
      ];

      let org = this.$q.localStorage.getItem("org");
      worksheet.getCell("A1").value = "No."; //ส่วนหัว
       worksheet.getCell("A1").font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
        worksheet.getCell("A1").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        worksheet.getCell("A1").alignment = {
          horizontal: "center",
          vertical: "middle",
        };

        worksheet.getCell("A2").value = ""; //ส่วนหัว
        worksheet.getCell("A2").font = {
          name: "Angsana New",
          family: 4,
          size: 20,
          bold: true,
        };
        worksheet.getCell("A2").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        worksheet.getCell("A2").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

        worksheet.getCell("B1").value = "Current"; //ส่วนหัว
         worksheet.getCell("B1").font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
        worksheet.getCell("B1").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        worksheet.getCell("B1").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

        worksheet.getCell("B2").value = "Week";
        worksheet.mergeCells("B2:C2");
        worksheet.getCell("B2").alignment = {
          horizontal: "center",
          vertical: "middle",
        };
        worksheet.getCell("B2").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };

        worksheet.getCell("C1").value = 14; //ส่วนหัว
        worksheet.getCell("C1").font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
        worksheet.getCell("C1").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };

        worksheet.getCell("D1").value = "24-Jan"; //ส่วนหัว
         worksheet.getCell("D1").font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
        worksheet.getCell("D1").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        worksheet.getCell("D1").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };

        worksheet.getCell("D2").value = "Period";
        worksheet.mergeCells("D2:E2");
        worksheet.getCell("D2").alignment = {
          horizontal: "center",
          vertical: "middle",
        };
        worksheet.getCell("D2").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };

        worksheet.getCell("E1").value = "29-Jan"; //ส่วนหัว
         worksheet.getCell("E1").font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };
        worksheet.getCell("E1").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
         worksheet.getCell("E1").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

        worksheet.getCell("F1").value = "Total Width Loss (Yds) L";
        worksheet.mergeCells("F1:F2");
        worksheet.getCell("F1").alignment = {
          horizontal: "center",
          vertical: "middle",
        };
        worksheet.getCell("F1").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
         worksheet.getCell("F1").font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 14,
        bold: false
      };

        worksheet.getCell("G1").value = "Average Observe Width (Inch) F";
        worksheet.mergeCells("G1:G2");
        worksheet.getCell("G1").alignment = {
          horizontal: "center",
          vertical: "middle",
        };
        worksheet.getCell("G1").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };

        worksheet.getCell("H1").value = "Total Mark J";
        worksheet.mergeCells("H1:H2");
        worksheet.getCell("H2").alignment = {
          horizontal: "center",
          vertical: "middle",
        };
        worksheet.getCell("H2").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };

        worksheet.getCell("I1").value = "% Width Loss";
        worksheet.mergeCells("I1:I2");
        worksheet.getCell("I1").alignment = {
          horizontal: "center",
          vertical: "middle",
        };
        worksheet.getCell("I1").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };


      

        worksheet.getCell("J1").value = "% Width Loss Target";
        worksheet.mergeCells("J1:J2");
        worksheet.getCell("J1").alignment = {
          horizontal: "center",
          vertical: "middle",
        };
        worksheet.getCell("J1").border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };


  if(this.start_date_for_graph == "" && this.end_date_for_graph == ""){

        var myNewDate = new Date(this.start);
        
        var myNewDate2 = new Date(myNewDate);

        var count = 0;
     for (var a1 = 1; a1 < 53; a1++) {
        
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
       }else{
          
         this.next_date = myNewDate.setDate(myNewDate.getDate() + 2);
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

  }else{
   
       var startDate = moment(this.start_date_for_graph,'DD-MM-YYYY')
       var endDate = moment(this.end_date_for_graph,'DD-MM-YYYY')

       var date_name_start = moment(startDate).format('dddd'); ;
       var date_name_end = moment(endDate).format('dddd');

    
      if(date_name_start != "Monday"){
        this.$q.notify({
          message: "Please Chose Start Date Monday Only" ,
          position:"center",
          color: "red-9",
        });
      }else if(date_name_end != "Saturday"){
        this.$q.notify({
          message: "Please Chose End Date Saturday Only" ,
          position:"center",
          color: "red-9",
        });
      }else{

       var first_date = new Date(startDate)

       var sec_date = new Date(this.end_date_for_graph)
       
       this.fix_date = endDate.diff(startDate, 'days')  
       this.fix_week = this.fix_date/7

      
       this.fix_week = Math.round(this.fix_week)
    
       if(this.fix_date < 7){
         this.fix_week = 1
       }
        //round
      console.log(this.fix_date)
      console.log(this.fix_week)
     
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
       
      this.date_use_data2.push({
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

       this.date_use_data2.push({
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
       this.date_use_data2.push({
         first_date : this.firstDate,
         sec_date : this.secDate,
         fmt_first_date:this.change_formar_first_date,
         fmt_sec_date:this.change_formar_sec_date
       })
       round++
    }
     
     }
     
      
      }
    
  }

 
        const params  = new FormData()
    

        for(var i = 0; i<this.date_use_data.length; i++){
        this.result_ttlmarker = 0
        this.total_widthloss = 0
        this.last_total_widthloss = 0
        this.total_width_inch = 0;
        this.total_width_inch =0;
        this.add_width_inch = 0;
        this.total_keep_target = 0;
        this.target_widthloss = 0;
        this.total_keep_targetx = 0;
        this.total_plug12 = 0;
        this.add_total_plug12 = 0;
        this.add_width_inch = 0
         this.first_width_inch = 0
         this.acc_act_widthloss = 0;
         this.acc_act_widthloss = 0;
        this.last_acc_act_widthloss = 0;
        this.result_ttlmarkerx = 0;
        this.total_widthlossx = 0;
        this.last_total_widthlossx = 0
        
        
        
    
          params.append("start",this.date_use_data[i].convert_axios1)
          params.append("end",this.date_use_data[i].convert_axios2)
          params.append("org",this.org)
         
        await axios({
          method:'post',
          url:this.$api_url + "/mainso.php/export_data_5",
          data:params
        }).then((resp)=>{
          
          this.all_data_arr=[]
         if (resp.data.data.length > 0) {
          resp.data.data.forEach((e) => {
            this.all_data_arr.push({
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
              splicelossper: parseFloat(e.SPLICE_LOSS_PER).toFixed(2),
              overlap: e.OVER_LAB_YD,
              overlapper: parseFloat(e.OVER_LAB_PER).toFixed(2),
              defect: parseFloat(e.DEFECT_YD).toFixed(2),
              defectper: parseFloat(e.DEFECT_PER).toFixed(2),
              totalcutout: parseFloat(e.TOTAL_CUTOUT_YD).toFixed(2),
              totalcutoutper: e.TOTAL_CUTOUT_PER,
              rement: parseFloat(e.REMENT_LOSS_YD).toFixed(2),
              rement_per: parseFloat(e.REMENT_LOSS_PER).toFixed(2),
              totallenthloss: parseFloat(
                e.TOTAL_LENGTH_LOSS_YD)
              .toFixed(2),
              totallenthlossper: parseFloat(
                e.TOTAL_LEN_LOSS_PER)
              .toFixed(2),
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
            });
            
             
          });
          
           this.all_ttl_marker = 0
           this.all_total_widthloss = 0
           this.last_all_total_widthloss = 0
            for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
              this.all_ttl_marker = Number(this.all_ttl_marker) + Number(this.all_data_arr[ax].ttl_marker)
              this.all_total_widthloss = (this.all_data_arr[ax].ttl_marker * this.all_data_arr[ax].widthloss)/100
              this.last_all_total_widthloss = Number(this.last_all_total_widthloss)+Number(this.all_total_widthloss) //F
            }
            this.all_ttl_marker_arr.push({
              value_ttl : this.all_ttl_marker
            })
            this.acc_act_widthlossx.push({
              value:this.last_all_total_widthloss
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
            this.all_ttl_marker_arr.push({
              value_ttl : 0
            })
             this.acc_act_widthlossx.push({
              value:0
            })
         } //if
    
        

          var x = i+3;
      
      worksheet.getCell("A"+x).value = i+1; //ส่วนหัว
      worksheet.getCell("A"+x).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
         worksheet.getCell("A"+x).alignment = {
        horizontal: "center",
        vertical: "middle",
      };

       /*  this.date_use_data.push({
            first_date:this.firstNewDate,
            sec_date:this.secNewDate,
            convert_axios1:this.convert_axios1,
            convert_axios2:this.convert_axios2
          }) */

      worksheet.getCell("B"+x).value = this.date_use_data[i].first_date + " - " + this.date_use_data[i].sec_date; //ส่วนหัว
      worksheet.getCell("B"+x).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
         worksheet.getCell("B"+x).alignment = {
        horizontal: "center",
        vertical: "middle",
      };


      worksheet.getCell("C"+x).value = i+1; //ส่วนหัว
      worksheet.getCell("C"+x).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
         worksheet.getCell("C"+x).alignment = {
        horizontal: "center",
        vertical: "middle",
      };



      worksheet.getCell("D"+x).value = this.date_use_data[i].first_date; //ส่วนหัว
      worksheet.getCell("D"+x).border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          
        };
         worksheet.getCell("D"+x).alignment = {
        horizontal: "center",
        vertical: "middle",
      };


        worksheet.getCell("E"+x).value = this.date_use_data[i].sec_date; //ส่วนหัว
      worksheet.getCell("E"+x).border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
         worksheet.getCell("E"+x).alignment = {
        horizontal: "center",
        vertical: "middle",
      };


       
     
        for(var j = 0; j<this.all_data_arr.length; j++){
         
          if(this.all_data_arr[j].widthloss !== ""){
         this.total_widthloss = (this.all_data_arr[j].ttl_marker * this.all_data_arr[j].widthloss)/100
         this.last_total_widthloss = Number(this.last_total_widthloss)+Number(this.total_widthloss)
          }
       
        }
       
         
         if(this.last_total_widthloss == "" || this.last_total_widthloss == null || this.last_total_widthloss == "NaN" || this.last_total_widthloss == 0){
        worksheet.getCell("F"+x).value = ""; //ส่วนหัว
      worksheet.getCell("F"+x).border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        }else{

      worksheet.getCell("F"+x).value = Number(parseFloat(this.last_total_widthloss).toFixed(2)); //ส่วนหัว
      worksheet.getCell("F"+x).border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
         worksheet.getCell("F"+x).alignment = {
        horizontal: "center",
        vertical: "middle",
      };

        } 

        
        for(var j = 0; j<this.all_data_arr.length; j++){
          this.result_ttlmarker = Number(this.result_ttlmarker) + Number(this.all_data_arr[j].ttl_marker)
        }
        
        

        
        for(var j = 0; j<this.all_data_arr.length; j++){
           if(this.all_data_arr[j].obs_width !== ""){
        this.total_width_inch = this.all_data_arr[j].obs_width * this.all_data_arr[j].ttl_marker;
        this.add_width_inch = Number(this.add_width_inch) + Number(this.total_width_inch);
           }
        }
        
        if(this.result_ttlmarker !== 0){
        this.add_width_inch = this.add_width_inch / this.result_ttlmarker;
        }
      
     

        if(this.add_width_inch == "" || this.add_width_inch == null || this.add_width_inch == "NaN" || this.add_width_inch == 0){ 
        worksheet.getCell("G"+x).value = ""; //ส่วนหัว
      worksheet.getCell("G"+x).border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        }else{
        worksheet.getCell("G"+x).value = Number(parseFloat(this.add_width_inch).toFixed(2)); //ส่วนหัว
      worksheet.getCell("G"+x).border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
         worksheet.getCell("G"+x).alignment = {
        horizontal: "center",
        vertical: "middle",
      };

        }

         if(this.result_ttlmarker == "" || this.result_ttlmarker == null || this.result_ttlmarker == "NaN"){
        worksheet.getCell("H"+x).value = ""; //ส่วนหัว
      worksheet.getCell("H"+x).border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        }else{
        worksheet.getCell("H"+x).value = Number(parseFloat(this.result_ttlmarker).toFixed(2)); //ส่วนหัว
      worksheet.getCell("H"+x).border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
         worksheet.getCell("H"+x).alignment = {
        horizontal: "center",
        vertical: "middle",
      };

        
        
        }

        
          for (var j = 0; j<this.all_data_arr.length; j++){
            this.target_widthloss = (this.all_data_arr[j].ttl_marker * this.all_data_arr[j].widthloss) /100;
          }
          if(this.target_widthloss > 0){
          this.keep_target_widthloss.push({
            target_wl : this.target_widthloss
          })
          }else{
            this.keep_target_widthloss.push({
              target_wl:""
            })
          }
        
         
         for(var k = 0; k<this.keep_target_widthloss.length; k++){
            if(this.keep_target_widthloss[k].target_wl !== ""){
                this.total_keep_target = Number(this.total_keep_target) + Number(this.keep_target_widthloss[k].target_wl)
            }
          } 
          if(this.result_ttlmarker !== 0){
          this.total_keep_targetx = (this.last_total_widthloss / this.result_ttlmarker)*100;
          this.total_keep_target_graph.push({
            value:this.total_keep_targetx,
            value_ax:k-1
          })
          }

  
          
          
         if(this.total_keep_targetx == "" || this.total_keep_targetx == null || this.total_keep_targetx == "NaN" || this.total_keep_targetx == 0){
        worksheet.getCell("I"+x).value = ""; //ส่วนหัว
      worksheet.getCell("I"+x).border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
        }else{
        worksheet.getCell("I"+x).value = this.total_keep_targetx/100 //ส่วนหัว
        worksheet.getCell("I"+x).numFmt = '0.00%'
      worksheet.getCell("I"+x).border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
         worksheet.getCell("I"+x).alignment = {
        horizontal: "center",
        vertical: "middle",
      };

        }
       
        
        this.keep_target.push({
          value:this.target
        })
        worksheet.getCell("J"+x).value = this.target/100 //ส่วนหัว
        worksheet.getCell("J"+x).numFmt = '0.00%'; 
      worksheet.getCell("J"+x).border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
         worksheet.getCell("J"+x).alignment = {
        horizontal: "center",
        vertical: "middle",
      };

        
      
        for (var axc = 0; axc < this.all_data_arr.length; axc++) {
         
         this.first_width_inch =   this.all_data_arr[axc].width_inch * this.all_data_arr[axc].ttl_marker;
          this.add_width_inch =
            Number(this.add_width_inch) + Number(this.first_width_inch);
        }
        
        if(this.result_ttlmarker !== 0){
        this.add_width_inch = this.add_width_inch / this.result_ttlmarker;
        }
      


         
        for (var ac = 0; ac < resp.data.data.length; ac++) {
          this.total_plug12 =
            Number(this.all_data_arr[ac].ttl_marker) *
            Number(this.all_data_arr[ac].plug12);
          this.add_total_plug12 =
            Number(this.add_total_plug12) + Number(this.total_plug12);
        }
        
        if(this.result_ttlmarker !== 0){
        this.add_total_plug12 = this.add_total_plug12 / this.result_ttlmarker;
        }


         for(var xyz1 = 2; xyz1<x+1; xyz1++){
        worksheet.getRow(xyz1).font = {
        name: 'Angsana New',
        color: { argb: 'FF000000' },
        family: 4,
        size: 11,
        bold: false
      };
         }

       
        
        })
        }

        

        this.sum_widthloss = 0
        this.sum_ttl_marker = 0
        this.sum_widthloss_ttl = 0
  
        

    
      
      this.$q.loading.hide({
          
        })
        
         
      /*     workbook.xlsx.writeBuffer().then((data) => {
        const blob = new Blob([data], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8",
        });
        saveAs(blob, "report_weekly_wla_data.xlsx");
      }); */  

      if(this.start_date_for_graph == "" && this.end_date_for_graph == ""){

     console.log(this.date_use_data)
     console.log(this.total_keep_target_graph)
     console.log(this.total_keep_target_graph.length)

     let total_keep_target_graph_minus = this.total_keep_target_graph.pop();

          for(var ax = 0; ax<this.total_keep_target_graph.length; ax++){
          
         this.date_data_for_graph.push({
           value:this.date_use_data[this.total_keep_target_graph[ax].value_ax].first_date+" - "+ this.date_use_data[this.total_keep_target_graph[ax].value_ax].sec_date
         })

         this.keep_target_for_graph.push({
           value:this.keep_target[this.total_keep_target_graph[ax].value_ax].value
         })

        
      } 
      console.log(this.total_keep_target_graph)

     
     

      let width_loss_ttl = []
      let index1 = 0;
      this.total_keep_target_graph.forEach(e => {
                width_loss_ttl[index1] = e.value
                index1++
            })

     let data_date = [];
            let index = 0;
            this.date_data_for_graph .forEach(e => {
                data_date[index] = e.value
                index++
            })
      
      let target = [];
            let index2 = 0;
            this.keep_target_for_graph .forEach(e => {
                target[index2] = e.value
                index2++
            })


  console.log(width_loss_ttl)
  console.log(data_date)
  console.log(target)
     
         

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
   boundaryGap: [0, '50%']
  },
  series: [
    {
      name:"actual",
      data: width_loss_ttl,
      type: 'line'
    },
    {
      name: "target",
      data: target,
      type: 'line'
    }, 
  /*    {
      data: acc_act_width,
      type: 'line'
    },  */
   
  ]
};

option && myChart.setOption(option);
      }else{
       console.log(this.date_use_data2)
        for(var ax = 0; ax<this.date_use_data2.length; ax++){
          const params = new FormData();
          
          params.append("start",this.date_use_data2[ax].first_date)
          params.append("end",this.date_use_data2[ax].sec_date)
          params.append("org",this.org)
         
        await axios({
          method:'post',
          url:this.$api_url + "/mainso.php/export_data_5",
          data:params
        }).then((resp)=>{
          console.log(resp.data)
          this.all_data_arr2=[]
         if (resp.data.data.length > 0) {
          resp.data.data.forEach((e) => {
            this.all_data_arr2.push({
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
              splicelossper: parseFloat(e.SPLICE_LOSS_PER).toFixed(2),
              overlap: e.OVER_LAB_YD,
              overlapper: parseFloat(e.OVER_LAB_PER).toFixed(2),
              defect: parseFloat(e.DEFECT_YD).toFixed(2),
              defectper: parseFloat(e.DEFECT_PER).toFixed(2),
              totalcutout: parseFloat(e.TOTAL_CUTOUT_YD).toFixed(2),
              totalcutoutper: e.TOTAL_CUTOUT_PER,
              rement: parseFloat(e.REMENT_LOSS_YD).toFixed(2),
              rement_per: parseFloat(e.REMENT_LOSS_PER).toFixed(2),
              totallenthloss: parseFloat(
                e.TOTAL_LENGTH_LOSS_YD)
              .toFixed(2),
              totallenthlossper: parseFloat(
                e.TOTAL_LEN_LOSS_PER)
              .toFixed(2),
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
            });
            
             
          });
          
           this.all_ttl_marker = 0
           this.all_total_widthloss = 0
           this.last_all_total_widthloss = 0
            for(var ax = 0; ax<this.all_data_arr.length; ax++){
              
              this.all_ttl_marker = Number(this.all_ttl_marker) + Number(this.all_data_arr[ax].ttl_marker)
              this.all_total_widthloss = (this.all_data_arr[ax].ttl_marker * this.all_data_arr[ax].widthloss)/100
              this.last_all_total_widthloss = Number(this.last_all_total_widthloss)+Number(this.all_total_widthloss) //F
            }
            this.all_ttl_marker_arr.push({
              value_ttl : this.all_ttl_marker
            })
            this.acc_act_widthlossx.push({
              value:this.last_all_total_widthloss
            })
         }else{
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
            this.all_ttl_marker_arr.push({
              value_ttl : 0
            })
             this.acc_act_widthlossx.push({
              value:0
            })
         }
        })
        this.result_ttlmarker2 = 0
        this.total_widthloss2 = 0
        this.last_total_widthloss2 = 0
        for(var ay = 0; ay<this.all_data_arr2.length; ay++){
          this.total_widthloss2 = (this.all_data_arr2[ay].ttl_marker * this.all_data_arr2[ay].widthloss)/100
          
          this.last_total_widthloss2 = Number(this.last_total_widthloss2) + Number(this.total_widthloss2)

          this.result_ttlmarker2 = Number(this.result_ttlmarker2) + Number(this.all_data_arr2[ay].ttl_marker)

      
      
        }

          if(this.result_ttlmarker2 !== 0){
          this.total_keep_targetx2 = (this.last_total_widthloss2 / this.result_ttlmarker2)*100;
          this.total_keep_target_graph2.push({
            value:this.total_keep_targetx2,
            value_ax:ax
          })
          }
       
     
         this.keep_target2.push({
        value:this.target
      })
       
        
      } 
      

        this.date_data_for_graph2 = []
             for(var ax = 0; ax<this.total_keep_target_graph2.length; ax++){
          console.log(this.total_keep_target_graph2[ax].value_ax)
         this.date_data_for_graph2.push({
           value:this.date_use_data2[this.total_keep_target_graph2[ax].value_ax].first_date+" - "+ this.date_use_data2[this.total_keep_target_graph2[ax].value_ax].sec_date
         })

         this.keep_target_for_graph2.push({
           value:this.keep_target2[this.total_keep_target_graph2[ax].value_ax].value
         })

        }

        console.log(this.date_data_for_graph2)
        console.log(this.keep_target_for_graph2)
        console.log(this.keep_target2)

        let width_loss_ttl2 = []
        let indexb = 0;
      this.total_keep_target_graph2 .forEach(e => {
                width_loss_ttl2[indexb] = e.value
                indexb++
            })

     let data_date2 = [];
            let indexa = 0;
            this.date_data_for_graph2 .forEach(e => {
                data_date2[indexa] = e.value
                indexa++
            })
      
      let target2 = [];
            let indexc = 0;
            this.keep_target2 .forEach(e => {
                target2[indexc] = e.value
                indexc++
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
    data: data_date2
  },
  yAxis: {
    type: 'value',
   boundaryGap: [0, '50%']
  },
  series: [
    {
      name:"actual",
      data: width_loss_ttl2,
      type: 'line'
    },
    {
      name: "target",
      data: target2,
      type: 'line'
    }, 
  /*    {
      data: acc_act_width,
      type: 'line'
    },  */
   
  ]
};

option && myChart.setOption(option);
        
          
      }
      
    },
  },
 /*  watch: {
    start(val) {
      let day = val.substring(0, 2);
      let month = val.substring(3, 5);
      let year = val.substring(6, 10);
      this.start_date = year + "/" + month + "/" + day;
    },
    end(val) {
      let day = val.substring(0, 2);
      let month = val.substring(3, 5);
      let year = val.substring(6, 10);
      this.end_date = year + "/" + month + "/" + day;
    },
  }, */
};
</script>

<style>
.myclass {
  padding-top : 50px;
  padding-right: 40px;
  padding-bottom:40px;
  padding-left:50px;
}</style>
