<template>
  <div style="max-width: 100%">
    <div class="bg-blue-10 text-white">
      <q-toolbar>
        <q-btn flat round dense />

        <q-toolbar-title>Setting</q-toolbar-title>

        <q-btn icon="cancel" flat fixed-right color="white" v-close-popup />
      </q-toolbar>
    </div>
    <br />

    <div class="q-mx-lg">
      <div class="row">
        <div class="col-2">
          <q-input
            filled
            v-model="type"
            bg-color="teal-4"
            :input-style="{ color: 'white' }"
            input-class="text-center"
            dense
            :readonly="true"
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            filled
            v-model="text"
            item-aligned="center"
            label="Woven"
            dense
            :readonly="true"
          />
        </div>
        <div class="col-1"></div>

        <div class="col-2">
          <q-input
            filled
            v-model="text"
            label="Light And Middle Weight"
            dense
            :readonly="true"
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            filled
            v-model="text"
            label="Fleece"
            dense
            :readonly="true"
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            filled
            v-model="text"
            label="Speical"
            dense
            :readonly="true"
          />
        </div>
      </div>
      <br />
      <div class="row">
        <div class="col-2">
          <q-input
            filled
            bg-color="blue-10"
            :input-style="{ color: 'white' }"
            input-class="text-center"
            v-model="end_loss"
            dense
            :readonly="true"
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="end_loss_woven"
            input-class="text-center"
            type="number"
            label=""
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>

        <div class="col-1"></div>

        <div class="col-2">
          <q-input
            v-model="end_loss_light"
            input-class="text-center"
            type="number"
            label=""
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="end_loss_fleece"
            input-class="text-center"
            type="number"
            label=""
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="end_loss_special"
            input-class="text-center"
            type="number"
            label=""
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>
      </div>

      <br />
      <div class="row">
        <div class="col-2">
          <q-input
            filled
            bg-color="blue-10"
            :input-style="{ color: 'white' }"
            input-class="text-center"
            v-model="splice_loss"
            dense
            :readonly="true"
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="splice_loss_woven"
            input-class="text-center"
            type="number"
            label=""
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>

        <div class="col-1"></div>

        <div class="col-2">
          <q-input
            v-model="splice_loss_light"
            input-class="text-center"
            type="number"
            label=""
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="splice_loss_fleece"
            input-class="text-center"
            type="number"
            label=""
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="splice_loss_special"
            input-class="text-center"
            type="number"
            label=""
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>
      </div>
      <br />
      <div class="row">
        <div class="col-2">
          <q-input
            filled
            bg-color="blue-10"
            :input-style="{ color: 'white' }"
            input-class="text-center"
            v-model="cut_loss"
            dense
            :readonly="true"
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="cut_loss_woven"
            label=""
            input-class="text-center"
            type="number"
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>

        <div class="col-1"></div>

        <div class="col-2">
          <q-input
            v-model="cut_loss_light"
            label=""
            input-class="text-center"
            type="number"
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="cut_loss_fleece"
            label=""
            input-class="text-center"
            type="number"
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="cut_loss_special"
            label=""
            input-class="text-center"
            type="number"
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>
      </div>
      <br />
      <div class="row">
        <div class="col-2">
          <q-input
            filled
            bg-color="blue-10"
            :input-style="{ color: 'white' }"
            input-class="text-center"
            v-model="rem_loss"
            dense
            :readonly="true"
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="rem_loss_woven"
            input-class="text-center"
            label=""
            type="number"
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>

        <div class="col-1"></div>

        <div class="col-2">
          <q-input
            v-model="rem_loss_light"
            input-class="text-center"
            type="number"
            label=""
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="rem_loss_fleece"
            input-class="text-center"
            type="number"
            label=""
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="rem_loss_special"
            input-class="text-center"
            label=""
            type="number"
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>
      </div>
      <br />
      <div class="row">
        <div class="col-2">
          <q-input
            filled
            bg-color="blue-10"
            :input-style="{ color: 'white' }"
            input-class="text-center"
            v-model="total_loss"
            dense
            :readonly="true"
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="total_woven"
            input-class="text-center"
            type="number"
            label=""
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>

        <div class="col-1"></div>

        <div class="col-2">
          <q-input
            v-model="total_light"
            input-class="text-center"
            type="number"
            label=""
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="total_fleece"
            input-class="text-center"
            type="number"
            label=""
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="total_special"
            input-class="text-center"
            type="number"
            label=""
            :bg-color="isdisable ? 'blue-grey-1 ' : ''"
            :disable="isdisable"
            dense
          />
        </div>
      </div>

      <br />
      <div class="row">
        <div class="col-3">
          <q-input
            filled
            v-model="date_cum"
            bg-color="blue-10"
            :input-style="{ color: 'white' }"
            input-class="text-center"
            dense
            :readonly="true"
          />
        </div>
        <div class="col-1"></div>

        <div class="col-4">
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
        </div>
      </div>
      <br />
      <div class="row">
        <div class="col-5"></div>
        <q-btn
          size="md"
          dense
          class="q-px-xl q-py-xห"
          color="primary"
          label="Save"
          @click="Update_setting()"
        />
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
export default {
  data() {
    return {
      type: "type",
      end_loss: "End Loss",
      splice_loss: "Splice Loss",
      cut_loss: "Cut Out Loss",
      rem_loss: "Remnants Loss",
      total_loss: "Total Length Loss",
      date_cum: "Date start Cummulated 24 Week",

      end_loss_woven: "",
      end_loss_light: "",
      end_loss_fleece: "",
      end_loss_special: "",

      splice_loss_woven: "",
      splice_loss_light: "",
      splice_loss_fleece: "",
      splice_loss_special: "",

      cut_loss_woven: "",
      cut_loss_light: "",
      cut_loss_fleece: "",
      cut_loss_special: "",

      rem_loss_woven: "",
      rem_loss_light: "",
      rem_loss_fleece: "",
      rem_loss_special: "",

      total_woven: "",
      total_light: "",
      total_fleece: "",
      total_special: "",

      start: "",
      rowx: [],
      keep_end: [],
      keep_splice: [],
      keep_cut: [],
      keep_rem: [],
      update_value: [],
      username: [],
      isdisable: true,
      columex: [
        {
          name: "Woven",
          label: "Woven",
          align: "center",
          field: "Woven",
        },
        {
          name: "Light",
          label: "Light & Middle Weight",
          align: "center",
          field: "Light",
        },
        {
          name: "Fleece",
          label: "Fleece",
          align: "center",
          field: "Fleece",
        },
        {
          name: "Special",
          label: "Special",
          align: "center",
          field: "Special",
        },
      ],
    };
  },
  mounted() {
    let username = this.$q.localStorage.getItem("username");
    let status = this.$q.localStorage.getItem("status");
    if (status == "Y") {
      this.isdisable = false;
    }
    this.rowx = [];
    axios
      .get(this.$api_url + "/mainso.php/get_target")
      .then((resp) => {
        this.rowx = resp.data.data;

        this.end_loss_woven = parseFloat(this.rowx[0].END_LOSS_WOVEN).toFixed(
          2
        );
        this.end_loss_light = parseFloat(this.rowx[0].END_LOSS_LIGHT).toFixed(
          2
        );
        this.end_loss_fleece = parseFloat(this.rowx[0].END_LOSS_FLEECE).toFixed(
          2
        );
        this.end_loss_special = parseFloat(
          this.rowx[0].END_LOSS_SPECIAL
        ).toFixed(2);

        this.splice_loss_woven = parseFloat(
          this.rowx[0].SPLICE_LOSS_WOVEN
        ).toFixed(2);
        this.splice_loss_light = parseFloat(
          this.rowx[0].SPLICE_LOSS_LIGHT
        ).toFixed(2);
        this.splice_loss_fleece = parseFloat(
          this.rowx[0].SPLICE_LOSS_FLEECE
        ).toFixed(2);
        this.splice_loss_special = parseFloat(
          this.rowx[0].SPLICE_LOSS_SPECIAL
        ).toFixed(2);

        this.cut_loss_woven = parseFloat(this.rowx[0].CUT_LOSS_WOVEN).toFixed(
          2
        );
        this.cut_loss_light = parseFloat(this.rowx[0].CUT_LOSS_LIGHT).toFixed(
          2
        );
        this.cut_loss_fleece = parseFloat(this.rowx[0].CUT_LOSS_FLEECE).toFixed(
          2
        );
        this.cut_loss_special = parseFloat(
          this.rowx[0].CUT_LOSS_SPECIAL
        ).toFixed(2);

        this.rem_loss_woven = parseFloat(this.rowx[0].REM_LOSS_WOVEN).toFixed(
          2
        );
        this.rem_loss_light = parseFloat(this.rowx[0].REM_LOSS_LIGHT).toFixed(
          2
        );
        this.rem_loss_fleece = parseFloat(this.rowx[0].REM_LOSS_FLEECE).toFixed(
          2
        );
        this.rem_loss_special = parseFloat(
          this.rowx[0].REM_LOSS_SPECIAL
        ).toFixed(2);

        this.total_woven =
          Number(this.end_loss_woven) +
          Number(this.splice_loss_woven) +
          Number(this.cut_loss_woven) +
          Number(this.rem_loss_woven);

        this.total_light =
          Number(this.end_loss_light) +
          Number(this.splice_loss_light) +
          Number(this.cut_loss_light) +
          Number(this.rem_loss_light);

        this.total_fleece =
          Number(this.end_loss_fleece) +
          Number(this.splice_loss_fleece) +
          Number(this.cut_loss_fleece) +
          Number(this.rem_loss_fleece);

        this.total_special =
          Number(this.end_loss_special) +
          Number(this.splice_loss_special) +
          Number(this.cut_loss_special) +
          Number(this.rem_loss_special);

        this.total_woven = parseFloat(this.total_woven).toFixed(2);
        this.total_light = parseFloat(this.total_light).toFixed(2);
        this.total_fleece = parseFloat(this.total_fleece).toFixed(2);
        this.total_special = parseFloat(this.total_special).toFixed(2);
      })
      .catch((error) => {
        console.log(error);
      });

    /* function convert_date(val){
          let day = val.substring(0, 2);
          let month = val.substring(3, 5);
          let year = val.substring(6, 10);
          this.start_date = year +"/"+ month +"/"+ day
          return this.start_date
          
        }
 */
    axios
      .get(this.$api_url + "/mainso.php/get_date")
      .then((resp) => {
        this.start = resp.data.data[0].START_DATE;
      })
      .catch((error) => {
        console.log(error);
      });
  },
  methods: {
    Update_setting() {
      const params = new FormData();
      params.append("end_woven", this.end_loss_woven);
      params.append("end_light", this.end_loss_light);
      params.append("end_fleece", this.end_loss_fleece);
      params.append("end_special", this.end_loss_special);
      params.append("splice_woven", this.splice_loss_woven);
      params.append("splice_light", this.splice_loss_light);
      params.append("splice_fleece", this.splice_loss_fleece);
      params.append("splice_special", this.splice_loss_special);
      params.append("cut_woven", this.cut_loss_woven);
      params.append("cut_light", this.cut_loss_light);
      params.append("cut_fleece", this.cut_loss_fleece);
      params.append("cut_special", this.cut_loss_special);
      params.append("rem_woven", this.rem_loss_woven);
      params.append("rem_light", this.rem_loss_light);
      params.append("rem_fleece", this.rem_loss_fleece);
      params.append("rem_special", this.rem_loss_special);
      let username = this.$q.localStorage.getItem("username");
      params.append("username", username);

      for (var pair of params.entries()) {
        console.log(pair[0] + ", " + pair[1]);
      }

      axios({
        method: "post",
        url: this.$api_url + "/mainso.php/update_target",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
        })
        .catch((error) => {
          console.log(error);
        });

      const params2 = new FormData();
      params2.append("start", this.start);
      params2.append("username", username);
      for (var pair of params2.entries()) {
        console.log(pair[0] + ", " + pair[1]);
      }

      axios({
        method: "post",
        url: this.$api_url + "/mainso.php/update_start_date",
        data: params2,
      })
        .then((resp) => {
          console.log(resp.data);
        })
        .catch((error) => {
          console.log(error);
        });

      axios
        .get(this.$api_url + "/mainso.php/get_target")
        .then((resp) => {
          this.rowx = resp.data.data;
          this.end_loss_woven = parseFloat(this.rowx[0].END_LOSS_WOVEN).toFixed(
            2
          );
          this.end_loss_light = parseFloat(this.rowx[0].END_LOSS_LIGHT).toFixed(
            2
          );
          this.end_loss_fleece = parseFloat(
            this.rowx[0].END_LOSS_FLEECE
          ).toFixed(2);
          this.end_loss_special = parseFloat(
            this.rowx[0].END_LOSS_SPECIAL
          ).toFixed(2);

          this.splice_loss_woven = parseFloat(
            this.rowx[0].SPLICE_LOSS_WOVEN
          ).toFixed(2);
          this.splice_loss_light = parseFloat(
            this.rowx[0].SPLICE_LOSS_LIGHT
          ).toFixed(2);
          this.splice_loss_fleece = parseFloat(
            this.rowx[0].SPLICE_LOSS_FLEECE
          ).toFixed(2);
          this.splice_loss_special = parseFloat(
            this.rowx[0].SPLICE_LOSS_SPECIAL
          ).toFixed(2);

          this.cut_loss_woven = parseFloat(this.rowx[0].CUT_LOSS_WOVEN).toFixed(
            2
          );
          this.cut_loss_light = parseFloat(this.rowx[0].CUT_LOSS_LIGHT).toFixed(
            2
          );
          this.cut_loss_fleece = parseFloat(
            this.rowx[0].CUT_LOSS_FLEECE
          ).toFixed(2);
          this.cut_loss_special = parseFloat(
            this.rowx[0].CUT_LOSS_SPECIAL
          ).toFixed(2);

          this.rem_loss_woven = parseFloat(this.rowx[0].REM_LOSS_WOVEN).toFixed(
            2
          );
          this.rem_loss_light = parseFloat(this.rowx[0].REM_LOSS_LIGHT).toFixed(
            2
          );
          this.rem_loss_fleece = parseFloat(
            this.rowx[0].REM_LOSS_FLEECE
          ).toFixed(2);
          this.rem_loss_special = parseFloat(
            this.rowx[0].REM_LOSS_SPECIAL
          ).toFixed(2);

          this.total_woven =
            Number(this.end_loss_woven) +
            Number(this.splice_loss_woven) +
            Number(this.cut_loss_woven) +
            Number(this.rem_loss_woven);

          this.total_light =
            Number(this.end_loss_light) +
            Number(this.splice_loss_light) +
            Number(this.cut_loss_light) +
            Number(this.rem_loss_light);

          this.total_fleece =
            Number(this.end_loss_fleece) +
            Number(this.splice_loss_fleece) +
            Number(this.cut_loss_fleece) +
            Number(this.rem_loss_fleece);

          this.total_special =
            Number(this.end_loss_special) +
            Number(this.splice_loss_special) +
            Number(this.cut_loss_special) +
            Number(this.rem_loss_special);

          this.total_woven = parseFloat(this.total_woven).toFixed(2);
          this.total_light = parseFloat(this.total_light).toFixed(2);
          this.total_fleece = parseFloat(this.total_fleece).toFixed(2);
          this.total_special = parseFloat(this.total_special).toFixed(2);
        })
        .catch((error) => {
          console.log(error);
        });

      axios
        .get(this.$api_url + "/mainso.php/get_date")
        .then((resp) => {
          this.start = resp.data.data[0].START_DATE;
        })
        .catch((error) => {
          console.log(error);
        });
    },
  },
};
</script>

<style>
.myclass {
  padding-top : 50px;
  padding-right: 40px;
  padding-bottom:40px;
  padding-left:50px;
}</style>
