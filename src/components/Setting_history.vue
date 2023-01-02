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
        <div class="col-1">
          <q-input
            filled
            v-model="type"
            bg-color="teal-4"
            :input-style="{ color: 'white' }"
            input-class="text-center"
            dense
         
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            filled
            v-model="this.year[0]"
            item-aligned="center"
            input-class="text-center"
            dense
           
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            filled
            v-model="this.year[1]"
            input-class="text-center"
            dense
       
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            filled
            v-model="this.year[2]"
            input-class="text-center"
            dense
            
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            filled
            v-model="this.year[3]"
            input-class="text-center"
            dense
      
          />
        </div>

        <div class="col-1"></div>

         <div class="col-1">
          <q-input
            filled
            v-model="this.year[4]"
            input-class="text-center"
            dense
           
          />
        </div>
      </div>
      <br />
      <div class="row">
        <div class="col-1">
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
            v-model="this.end_loss_value[0]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.end_loss_value[1]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.end_loss_value[2]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.end_loss_value[3]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        <div class="col-1"></div>
        <div class="col-1">
          <q-input
            v-model="this.end_loss_value[4]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
      </div>

      <br />
      <div class="row">
        <div class="col-1">
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
            v-model="this.splice_loss_value[0]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
           v-model="this.splice_loss_value[1]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.splice_loss_value[2]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.splice_loss_value[3]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.splice_loss_value[4]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
      </div>
      <br />
      <div class="row">
        <div class="col-1">
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
            v-model="this.cut_out_loss_value[0]"
            label=""
            input-class="text-center"
            type="number"
            dense
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.cut_out_loss_value[1]"
            label=""
            input-class="text-center"
            type="number"
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.cut_out_loss_value[2]"
            label=""
            input-class="text-center"
            type="number"
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.cut_out_loss_value[3]"
            label=""
            input-class="text-center"
            type="number"
            dense
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.cut_out_loss_value[4]"
            label=""
            input-class="text-center"
            type="number"
            dense
          />
        </div>
      </div>
      <br />
      <div class="row">
        <div class="col-1">
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
            v-model="this.remnants_loss_value[0]"
            input-class="text-center"
            label=""
            type="number"
            dense
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.remnants_loss_value[1]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.remnants_loss_value[2]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.remnants_loss_value[3]"
            input-class="text-center"
            label=""
            type="number"
            dense
          />
        </div>

         <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.remnants_loss_value[4]"
            input-class="text-center"
            label=""
            type="number"
            dense
          />
        </div>

        
      </div>
      <br />
      <div class="row">
        <div class="col-1">
          <q-input
            filled
            bg-color="blue-10"
            :input-style="{ color: 'white' }"
            input-class="text-center"
            v-model="total_spread_loss"
            dense
            :readonly="true"
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.total_spread_loss_per[0]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.total_spread_loss_per[1]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.total_spread_loss_per[2]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.total_spread_loss_per[3]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.total_spread_loss_per[4]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
      </div>

      <br />
      <div class="row">
        <div class="col-1">
          <q-input
            filled
            bg-color="blue-10"
            :input-style="{ color: 'white' }"
            input-class="text-center"
            v-model="total_width_loss"
            dense
            :readonly="true"
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.total_width_loss_per[0]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.total_width_loss_per[1]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.total_width_loss_per[2]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.total_width_loss_per[3]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.total_width_loss_per[4]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
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
      type: "YEAR",
      end_loss: "End Loss",
      splice_loss: "Splice Loss",
      cut_loss: "Cut Out Loss",
      rem_loss: "Remnants Loss",
      total_loss: "Total Length Loss",
      total_spread_loss :"Total Spread Loss Per",
      total_width_loss :"Total Width Loss per",
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
      year:[],
      end_loss_value:[],
      splice_loss_value:[],
      cut_out_loss_value:[],
      remnants_loss_value:[],
      total_spread_loss_per:[],
      total_width_loss_per:[],
      
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
    this.org = this.$q.localStorage.getItem("org");
    let status = this.$q.localStorage.getItem("status");

   

      const params3 = new FormData();
      params3.append("org",this.org)
      axios({
        method: "post",
        url: this.$api_url + "/mainso.php/check_history_per",
        data: params3,
      }).then((resp)=>{
        console.log(resp.data)
         var i = 0;

              resp.data.data.forEach((e) => {
                this.year[i] = e.YEAR;
                this.end_loss_value[i] = parseFloat(e.END_LOSS).toFixed(2);
                this.splice_loss_value[i] = parseFloat(e.SPLICE_LOSS).toFixed(2),
                this.cut_out_loss_value[i] = parseFloat(e.CUT_OUT_LOSS).toFixed(2),
                this.remnants_loss_value[i] = parseFloat(e.REMNANTS_LOSS).toFixed(2)
                this.total_spread_loss_per[i] = parseFloat(e.TOTAL_SPREAD_LOSS_PER).toFixed(2)
                this.total_width_loss_per[i] = parseFloat(e.TOTAL_WIDTH_LOSS_PER).toFixed(2)
                i++;
              });
        
      })
  
      

      
  },
  methods: {
    Update_setting() {
      const params = new FormData();
      for(var ax = 0; ax<this.year.length; ax++){
        params.append("year",this.year[ax])
        params.append("end_loss",this.end_loss_value[ax])
        params.append("splice_loss",this.splice_loss_value[ax])
        params.append("cut_out_loss",this.cut_out_loss_value[ax])
        params.append("remnants_loss",this.remnants_loss_value[ax])
        params.append("total_spread_loss_per",this.total_spread_loss_per[ax])
        params.append("total_width_loss_per",this.total_width_loss_per[ax])
        params.append("index_no",ax+1)
        params.append("org",this.org)
     
      for (var pair of params.entries()) {
                  console.log(pair[0] + ", " + pair[1]);
      }
    
       axios({
        method: "post",
        url: this.$api_url + "/mainso.php/update_history_per",
        data: params,
      })
        .then((resp) => {
          if(resp.data.status == true){
            this.$q.notify({
          message: "Update Success" ,
          position:"center",
          color: "green-9",
        });
       
          }
        })
        .catch((error) => {
          console.log(error);
        });
 
  }

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
