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
      style="width: 200px"
      disable
      label="Chose Org"
    />

    <q-btn
      size="md"
      dense
      class="q-px-xl q-py-xs"
      color="primary"
      label="Export"
      @click="exportexcel()"
    />
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
      date_start: "02/26/2022",
      data_count: 10,
      all_date: [],
      rowexport: [],
      rowexport2: [],
      x: 0,
      end_loss_plug12: [],
      keep_ttlmarker: [],
      keep_avg_end: [],
      keep_splice_Loss: [],
      keep_cut_out_loss: [],
      keep_remnants: [],
      keep_total_loss: [],
      keep_plug12_inch: [],
      keep_line_l: [],
      keep_avg_loss_plug12: [],
      keep_avg_length_yd: [],
      keep_sum_length_splice_loss: [],
      keep_sum_cut_loss_out: [],
      keep_perplug12: [],
      keep_total_loss_per: [],
      keep_total_width_loss_yds: [],
    };
  },
  mounted() {
    let org = this.$q.localStorage.getItem("org");
    this.org = org;
  },
  methods: {
    checkdata() {
      for (var x = 0; x < this.multiple.length; x++) {
        
      }
    },
    async exportexcel() {
      this.rowexport = [];
      this.end_loss_plug12 = [];
      this.keep_ttlmarker = [];
      this.keep_avg_end = [];
      this.keep_splice_Loss = [];
      this.keep_cut_out_loss = [];
      this.keep_remnants = [];
      this.keep_total_loss = [];
      this.keep_plug12_inch = [];
      this.keep_line_l = [];
      this.keep_avg_loss_plug12 = [];
      this.keep_avg_length_yd = [];
      this.keep_sum_length_splice_loss = [];
      this.keep_sum_cut_loss_out = [];
      this.keep_perplug12 = [];
      this.keep_splice_loss_per = [];
      this.keep_cut_out_loss_per = [];
      this.keep_remnants_per = [];
      this.keep_total_loss_per = [];
      this.check_value = "";
      this.keep_actual = [];
      this.export_out = [];
      this.keep_total_width_loss_yds = [];

      var x = 0;
      const workbook = new Excel.Workbook();
      workbook.creator = "Nanyang";
      workbook.lastModifiedBy = "Admin";
      workbook.created = new Date(2021, 8, 30);
      workbook.modified = new Date();
      workbook.lastPrinted = new Date(2021, 7, 27);
      const worksheet = workbook.addWorksheet("New Sheet");
      const worksheetx = workbook.addWorksheet("Data");

      worksheet.columns = [
        { key: "A", width: 10 },
        { key: "B", width: 7.63 },
        { key: "C", width: 12 },
        { key: "D", width: 50 },
        { key: "E", width: 10 },
        { key: "F", width: 8.7 },
        { key: "G", width: 8.7 },
        { key: "H", width: 6.7 },
        { key: "I", width: 6.7 },
        { key: "J", width: 12 },
        { key: "K", width: 12 },
        { key: "L", width: 6.7 },
        { key: "M", width: 6.7 },
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
      worksheet.getCell("A1").value =
        "WIDTH LOSS AUDIT WEEKLY REPORT ( MU-WLA-1 ) : NY" + org; //ส่วนหัว

      worksheet.getCell("A1").font = {
        name: "Angsana New",
        color: { argb: "FF0120E5" },
        family: 4,
        size: 20,
        bold: true,
      };

      worksheet.getCell("A2").value = "Period :"; //ส่วนหัว
      worksheet.getCell("A2").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      worksheet.getCell("B2").value = this.start_date; //ส่วนหัว

      worksheet.getCell("C2").value = "to"; //ส่วนหัว
      worksheet.getCell("C2").alignment = {
        horizontal: "center",
        vertical: "middle",
      };

      worksheet.getCell("D2").value = this.end_date; //ส่วนหัว

      worksheet.getCell("A3").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      worksheet.getCell("B3").value = "";
      worksheet.mergeCells("B3:D3");
      worksheet.getCell("B3").alignment = {
        horizontal: "center",
        vertical: "middle",
      };
      worksheet.getCell("B3").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      worksheet.getCell("E3").value = "Fabric";
      worksheet.mergeCells("E3:F3");
      worksheet.getCell("E3").alignment = {
        horizontal: "center",
        vertical: "middle",
      };
      worksheet.getCell("E3").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      worksheet.getCell("G3").value = "Marker";
      worksheet.mergeCells("G3:J3");
      worksheet.getCell("G3").alignment = {
        horizontal: "center",
        vertical: "middle",
      };
      worksheet.getCell("G3").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      worksheet.getCell("K3").value = "Width Loss";
      worksheet.mergeCells("K3:M3");
      worksheet.getCell("K3").alignment = {
        horizontal: "center",
        vertical: "middle",
      };
      worksheet.getCell("K3").border = {
        top: { style: "double" },
        left: { style: "double" },
        bottom: { style: "double" },
        right: { style: "double" },
      };

      var a = 4; // first row head column Date per table
      var b = 5; //mergeA
      var count = 0;

      const params = new FormData();

      params.append("start", this.start_date);
      params.append("end", this.end_date);
      params.append("org", org);
      for (var pair of params.entries()) {
        console.log(pair[0] + ", " + pair[1]);
      }

      await axios({
        method: "post",
        url: this.$api_url + "/mainso.php/export_data_4",
        data: params,
      }).then((resp) => {
        
        this.rowexport = [];

        function check_null(val) {
          if (
            val == "" ||
            val == null ||
            val == "NaN" ||
            val == "null" ||
            val == undefined
          ) {
            return 0;
          } else {
            return val;
          }
        }

        if (resp.data.data.length > 0) {
          resp.data.data.forEach((e) => {
            this.rowexport.push({
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
              remnants: check_null(e.REMNANTS_WEIGHT),
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
              sectionlosslenth: check_null(e.SECTION_LOSS_YD),
              sectionlossper: check_null(e.SECTION_LOSS_PER),
              spliceloss: check_null(e.SPLICE_LOSS_YD),
              splicelossper: parseFloat(check_null(e.SPLICE_LOSS_PER)).toFixed(
                2
              ),
              overlap: e.OVER_LAB_YD,
              overlapper: parseFloat(e.OVER_LAB_PER).toFixed(2),
              defect: parseFloat(e.DEFECT_YD).toFixed(2),
              defectper: parseFloat(e.DEFECT_PER).toFixed(2),
              totalcutout: parseFloat(check_null(e.TOTAL_CUTOUT_YD)).toFixed(2),
              totalcutoutper: check_null(e.TOTAL_CUTOUT_PER),
              rement: parseFloat(check_null(e.REMENT_LOSS_YD)).toFixed(2),
              rement_per: parseFloat(check_null(e.REMENT_LOSS_PER)).toFixed(2),
              totallenthloss: parseFloat(
                check_null(e.TOTAL_LENGTH_LOSS_YD)
              ).toFixed(2),
              totallenthlossper: parseFloat(
                check_null(e.TOTAL_LEN_LOSS_PER)
              ).toFixed(2),
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

          worksheet.getCell("A" + a).value = "Date";
          worksheet.mergeCells("A" + a + ":" + "A" + b);
          worksheet.getCell("A" + a).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("A" + a).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
          this.length = resp.data.data.length + a + 2;

          for (var ab = a + 2; ab < this.length; ab++) {
            var y = "A" + ab;

            var z = ab - a - 2;
            worksheet.getCell(y).value = this.rowexport[z].mu_date;
            worksheet.getCell(y).alignment = { horizontal: "center" };
            worksheet.getCell(y).border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }
          var ty = this.length;

          worksheet.getCell("A" + ty).value = "Weekly Total";
          worksheet.mergeCells("A" + ty + ":" + "I" + ty);
          worksheet.getCell("A" + ty).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("A" + ty).border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
          this.result_ttlmarker = 0;
          for (
            var re_ttlmarker = 0;
            re_ttlmarker < resp.data.data.length;
            re_ttlmarker++
          ) {
            this.result_ttlmarker =
              Number(this.result_ttlmarker) +
              Number(this.rowexport[re_ttlmarker].ttl_marker);
          }
          this.keep_ttlmarker.push({
            result_ttlmarker: this.result_ttlmarker,
          });
          worksheet.getCell("J" + ty).value = Number(
            this.result_ttlmarker
          )
          worksheet.getCell("J"+ty).numFmt = '#,##0.00';
          worksheet.getCell("J" + ty).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("J" + ty).border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };

          this.result_plug12 = 0;
          for (
            var re_plug1_2 = 0;
            re_plug1_2 < resp.data.data.length;
            re_plug1_2++
          ) {
            this.result_plug12 =
              Number(this.result_plug12) +
              Number(this.rowexport[re_plug1_2].plug12);
          }

          this.result_plug12 = this.result_plug12 / resp.data.data.length;
          this.keep_plug12_inch.push({
            result_plug12_inch: this.result_plug12,
          });

          worksheet.getCell("K" + ty).value = "";
          worksheet.getCell("K" + ty).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("K" + ty).border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };

          this.result_perplug12 = 0;
          for (
            var re_perplug12 = 0;
            re_perplug12 < resp.data.data.length;
            re_perplug12++
          ) {
            this.result_perplug12 =
              Number(this.result_perplug12) +
              Number(this.rowexport[re_perplug12].widthloss);
          }
          this.keep_perplug12.push({
            result_perplug12: this.result_perplug12,
          });

          //C------------------------------------------------------------------------
          worksheet.getCell("B" + a).value = "S/O no.";
          worksheet.mergeCells("B" + a + ":" + "B" + b);
          worksheet.getCell("B" + a).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("B" + a).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
          this.length = resp.data.data.length + a + 2;

          for (var ab = a + 2; ab < this.length; ab++) {
            var y = "B" + ab;

            var z = ab - a - 2;
            worksheet.getCell(y).value = Number(this.rowexport[z].gm_no_short);
            worksheet.getCell(y).alignment = { horizontal: "center" };
            worksheet.getCell(y).border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }

          //E-----------------------------------------------------------
          worksheet.getCell("C" + a).value = "GM08 no.";
          worksheet.mergeCells("C" + a + ":" + "C" + b);
          worksheet.getCell("C" + a).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("C" + a).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
          this.length = resp.data.data.length + a + 2;

          for (var ab = a + 2; ab < this.length; ab++) {
            var y = "C" + ab;

            var z = ab - a - 2;
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

          worksheet.getCell("D" + a).value = "Fabric Type";
          worksheet.mergeCells("D" + a + ":" + "D" + b);
          worksheet.getCell("D" + a).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("D" + a).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
          this.length = resp.data.data.length + a + 2;

          for (var ab = a + 2; ab < this.length; ab++) {
            var y = "D" + ab;

            var z = ab - a - 2;

            worksheet.getCell(y).value = this.rowexport[z].fabric_type;
            worksheet.getCell(y).alignment = { horizontal: "center" };
            worksheet.getCell(y).border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }
          worksheet.getCell("E" + a).value = "FT";
          worksheet.mergeCells("E" + a + ":" + "E" + b);
          worksheet.getCell("E" + a).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("E" + a).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
          this.length = resp.data.data.length + a + 2;

          for (var ab = a + 2; ab < this.length; ab++) {
            var y = "E" + ab;

            var z = ab - a - 2;

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

          worksheet.getCell("F" + a).value = "Observe";
          worksheet.getCell("F" + a).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("F" + a).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
          worksheet.getCell("F" + b).value = "Width";
          worksheet.getCell("F" + b).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("F" + b).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
          this.length = resp.data.data.length + a + 2;

          for (var ab = a + 2; ab < this.length; ab++) {
            var y = "F" + ab;

            var z = ab - a - 2;
            worksheet.getCell(y).value = Number(this.rowexport[z].obs_width)
            worksheet.getCell(y).numFmt = '0.00';
            worksheet.getCell(y).alignment = { horizontal: "center" };
            worksheet.getCell(y).border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }

          //H-------------------------------------------------------
          worksheet.getCell("G" + a).value = "Width";
          worksheet.getCell("G" + a).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("G" + a).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
          worksheet.getCell("G" + b).value = "(inch)";
          worksheet.getCell("G" + b).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("G" + b).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
          this.length = resp.data.data.length + a + 2;

          for (var ab = a + 2; ab < this.length; ab++) {
            var y = "G" + ab;

            var z = ab - a - 2;
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

          worksheet.getCell("I" + a).value = "Length";
          worksheet.getCell("I" + a).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("I" + a).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
          worksheet.getCell("I" + b).value = "(yd)";
          worksheet.getCell("I" + b).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("I" + b).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };

          this.length = resp.data.data.length + b; //2 มาจากไหน?

          for (var ab = a + 2; ab <= this.length; ab++) {
            var y = "I" + ab;

            var z = ab - a - 2;
            worksheet.getCell(y).value = Number(
              this.rowexport[z].length_ydb
            )
            worksheet.getCell(y).numFmt = '0.00';
            worksheet.getCell(y).alignment = { horizontal: "center" };
            worksheet.getCell(y).border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }
          worksheet.getCell("H" + a).value = "Marker VS";
          worksheet.getCell("H" + a).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("H" + a).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
          worksheet.getCell("H" + b).value = "obs.width";
          worksheet.getCell("H" + b).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("H" + b).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };

          this.length = resp.data.data.length + b; //2 มาจากไหน?

          for (var ab = a + 2; ab <= this.length; ab++) {
            var y = "H" + ab;

            var z = ab - a - 2;

            this.mark_obs =
              Number(this.rowexport[z].obs_width) -
              Number(this.rowexport[z].width_inch);
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
          worksheet.getCell("J" + a).value = "Ttl Marker";
          worksheet.getCell("J" + a).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("J" + a).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
          worksheet.getCell("J" + b).value = "Length(yd)";
          worksheet.getCell("J" + b).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("J" + b).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
          this.length = resp.data.data.length + a + 2;

          for (var ab = a + 2; ab < this.length; ab++) {
            var y = "J" + ab;

            var z = ab - a - 2;
            worksheet.getCell(y).value = Number(
              this.rowexport[z].ttl_marker
            )
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
          worksheet.getCell("K" + a).value = "Plug1+2";
          worksheet.getCell("K" + a).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("K" + a).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
          worksheet.getCell("K" + b).value = "(inch)";
          worksheet.getCell("K" + b).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("K" + b).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
          this.length = resp.data.data.length + a + 2;

          for (var ab = a + 2; ab < this.length; ab++) {
            var y = "K" + ab;

            var z = ab - a - 2;
            worksheet.getCell(y).value = Number(
              check_null(this.rowexport[z].plug12)
            )
            worksheet.getCell(y).numFmt = '0.00';
            worksheet.getCell(y).alignment = { horizontal: "center" };
            worksheet.getCell(y).border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }
          worksheet.getCell("L" + a).value = "Width Loss";
          worksheet.getCell("L" + a).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("L" + a).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
          worksheet.getCell("L" + b).value = "(YDS)";
          worksheet.getCell("L" + b).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("L" + b).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };

          this.length = resp.data.data.length + b; //2 มาจากไหน?

          for (var ab = a + 2; ab <= this.length; ab++) {
            var y = "L" + ab;

            var z = ab - a - 2;
            this.width_loss_yds = 0;

            this.width_loss_yds =
              (this.rowexport[z].ttl_marker * this.rowexport[z].widthloss) /
              100;
            this.keep_total_width_loss_yds.push({
              total_yds: this.width_loss_yds,
            });
            worksheet.getCell(y).value = Number(
              this.width_loss_yds
            )
            worksheet.getCell(y).numFmt = '0.00';
            worksheet.getCell(y).alignment = { horizontal: "center" };
            worksheet.getCell(y).border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          }

          this.total_widthloss_yds = 0;
          for (var axx = 0; axx < resp.data.data.length; axx++) {
            this.total_widthloss_yds =
              this.total_widthloss_yds +
              this.keep_total_width_loss_yds[axx].total_yds;
          }
          worksheet.getCell("L" + ty).value = Number(
            this.total_widthloss_yds
          )
          worksheet.getCell("L" + ty).numFmt = '0.00';
          worksheet.getCell("L" + ty).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("L" + ty).border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };

          worksheet.getCell("M" + a).value = "%";
          worksheet.getCell("M" + a).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("M" + a).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };
          worksheet.getCell("M" + b).value = 1.50/100;
          worksheet.getCell("M" +b).numFmt = '0.00%';
          worksheet.getCell("M" + b).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("M" + b).border = {
            top: { style: "double" },
            left: { style: "double" },
            bottom: { style: "double" },
            right: { style: "double" },
          };

          this.length = resp.data.data.length + b; //2 มาจากไหน?

          for (var ab = a + 2; ab <= this.length; ab++) {
            var y = "M" + ab;

            var z = ab - a - 2;
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
          worksheet.getCell("M" + ty).value = "";
          worksheet.getCell("M" + ty).alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          worksheet.getCell("M" + ty).border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        }

        var c = this.length + 2;
        var d = this.length + 2;
        var e = this.length + 5;
        var f = this.length + 6;
        var g = this.length + 7;
        var h = this.length + 8;
        var i = this.length + 9;
        var j = this.length + 10;
        var k = this.length + 11;
        var l = this.length + 12; //endloss
        var m = this.length + 13; //splice loss
        var n = this.length + 14; //Cut - out Loss
        var o = this.length + 15; //Remnants Loss
        var p = this.length + 16; //Total Loss

        for (var xyz1 = 2; xyz1 < p; xyz1++) {
          worksheet.getRow(xyz1).font = {
            name: "Angsana New",
            color: { argb: "FF000000" },
            family: 4,
            size: 11,
            bold: false,
          };
        }

        let org = this.$q.localStorage.getItem("org");

        worksheet.getCell("A" + d).value = "Weekly Average";
        worksheet.mergeCells("A" + d + ":" + "E" + d);
        worksheet.getCell("A" + d).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
        worksheet.getCell("A" + d).border = {
          top: { style: "double" },
          left: { style: "double" },
          bottom: { style: "double" },
          right: { style: "double" },
        };
        this.result_sum_ttlmarker = 0;
        for (var jd = 0; jd < this.keep_ttlmarker.length; jd++) {
          this.result_sum_ttlmarker =
            Number(this.result_sum_ttlmarker) +
            Number(this.keep_ttlmarker[jd].result_ttlmarker);
         
        }
        this.total_obs = 0;
        this.add_total_obs = 0; //F*J
        for (var ax = 0; ax < resp.data.data.length; ax++) {
          this.total_obs =
            this.rowexport[ax].ttl_marker * this.rowexport[ax].obs_width;
          this.add_total_obs =
            Number(this.add_total_obs) + Number(this.total_obs);
        }

        this.total_obs_width_avg = this.add_total_obs / this.result_ttlmarker;
        worksheet.getCell("F" + d).value = Number(this.total_obs_width_avg)
        worksheet.getCell("F"+d).numFmt = '0.00';
        worksheet.getCell("F" + d).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
        worksheet.getCell("F" + d).border = {
          top: { style: "double" },
          left: { style: "double" },
          bottom: { style: "double" },
          right: { style: "double" },
        };
        this.total_width_inch = 0;
        this.add_width_inch = 0;
        for (var axc = 0; axc < resp.data.data.length; axc++) {
          this.total_width_inch =
            this.rowexport[axc].width_inch * this.rowexport[axc].ttl_marker;
          this.add_width_inch =
            Number(this.add_width_inch) + Number(this.total_width_inch);
        }
        this.add_width_inch = this.add_width_inch / this.result_ttlmarker;

        worksheet.getCell("G" + d).value = Number(
          this.add_width_inch)
          worksheet.getCell("G"+d).numFmt = '0.00';
        worksheet.mergeCells("G" + d + ":" + "J" + d);
        worksheet.getCell("G" + d).alignment = {
          horizontal: "left",
          vertical: "middle",
        };
        worksheet.getCell("G" + d).border = {
          top: { style: "double" },
          left: { style: "double" },
          bottom: { style: "double" },
          right: { style: "double" },
        };

        this.total_plug12 = 0;
        this.add_total_plug12 = 0;
        for (var ac = 0; ac < resp.data.data.length; ac++) {
          this.total_plug12 =
            Number(this.rowexport[ac].ttl_marker) *
            Number(this.rowexport[ac].plug12);
          this.add_total_plug12 =
            Number(this.add_total_plug12) + Number(this.total_plug12);
        }
        
        this.add_total_plug12 = this.add_total_plug12 / this.result_ttlmarker;

        worksheet.getCell("K" + d).value = Number(
          this.add_total_plug12
        )
        worksheet.getCell("K"+d).numFmt = '0.00';
        worksheet.getCell("K" + d).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
        worksheet.getCell("K" + d).border = {
          top: { style: "double" },
          left: { style: "double" },
          bottom: { style: "double" },
          right: { style: "double" },
        };

        worksheet.getCell("L" + d).value = "";
        worksheet.getCell("L" + d).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
        worksheet.getCell("L" + d).border = {
          top: { style: "double" },
          left: { style: "double" },
          bottom: { style: "double" },
          right: { style: "double" },
        };

        this.total_target = 0;
        this.total_target =
          (this.total_widthloss_yds / this.result_ttlmarker) * 100;
        worksheet.getCell("M" + d).value =this.total_target/100
        worksheet.getCell("M"+d).numFmt = '0.00%';
        worksheet.getCell("M" + d).alignment = {
          horizontal: "center",
          vertical: "middle",
        };
        worksheet.getCell("M" + d).border = {
          top: { style: "double" },
          left: { style: "double" },
          bottom: { style: "double" },
          right: { style: "double" },
        };

        
        
        workbook.xlsx.writeBuffer().then((data) => {
        const blob = new Blob([data], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8",
        });
        saveAs(blob, "Weekly Mu WLA-1.xlsx");
      }); 
          
      });
    },
  },
  watch: {
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
  },
};
</script>

<style></style>
