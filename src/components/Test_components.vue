<template>
 <div class="bg-blue-10 text-white">
      <q-toolbar>
        <q-btn flat round dense />

        <q-toolbar-title>Setting History Performance</q-toolbar-title>

        <q-btn icon="cancel" flat fixed-right color="white" v-close-popup />
      </q-toolbar>
    </div>
  <div class="q-pa-md">
    <q-carousel
      v-model="slide"
      transition-prev="slide-right"
      transition-next="slide-left"
      animated
      control-color="primary"
      class="rounded-borders"
      style="min-height: 500px"
    >
    
      <q-carousel-slide name="style">
      
      
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
      <br/>
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
  <br/>
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
      
      <br/>
      <div class="row">
        <div class="col-1">
          <q-input
            filled
            bg-color="blue-10"
            :input-style="{ color: 'white' }"
            input-class="text-center"
            v-model="cut_out_loss"
            dense
            :readonly="true"
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.cut_out_loss_value[0]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.cut_out_loss_value[1]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.cut_out_loss_value[2]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.cut_out_loss_value[3]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        <div class="col-1"></div>
        <div class="col-1">
          <q-input
            v-model="this.cut_out_loss_value[4]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        
      </div>

      <br/>
      <div class="row">
        <div class="col-1">
          <q-input
            filled
            bg-color="blue-10"
            :input-style="{ color: 'white' }"
            input-class="text-center"
            v-model="remnants_loss"
            dense
            :readonly="true"
          />
        </div>

        <div class="col-1"></div>

        <div class="col-1">
          <q-input
            v-model="this.remnants_loss_value[0]"
            input-class="text-center"
            type="number"
            label=""
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
            type="number"
            label=""
            dense
          />
        </div>
        <div class="col-1"></div>
        <div class="col-1">
          <q-input
            v-model="this.remnants_loss_value[4]"
            input-class="text-center"
            type="number"
            label=""
            dense
          />
        </div>
        
      </div>
        <br/>
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

       <br/>
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
      </q-carousel-slide>
      <q-carousel-slide name="tv"  >
        
        <div class="q-ma-xl  justify-center ">
         <div class="row q-ma-md justify-center">
          <div class="col-1"></div>
          <div class="col-3">
          <q-input
            filled
            v-model="year_insert"
            label="Year ***กรอกเป็น คริสศักราช"
            label-color="orange-9"
            :input-style="{ color: 'black' }"
            input-class="text-center"
            dense
         
          />
        </div>
        </div>
         <br>
          <div class="row q-ma-md justify-center">
          <div class="col-1"></div>
          <div class="col-3">
          <q-input
            filled
            v-model="end_loss_insert"
            label="End Loss"
            label-color="green-9"
            :input-style="{ color: 'black' }"
            input-class="text-center"
            dense
         
          />
        </div>
          <div class="col-1" ></div>
          <div class="col-3">
            <q-input
            filled
            v-model="splice_loss_insert"
            label="Splice Loss"
            label-color="green-9"
            :input-style="{ color: 'black' }"
            input-class="text-center"
            dense
         
          />
          </div>
          </div>
           <br>
           <div class="row q-ma-md justify-center">
          <div class="col-1"></div>
          <div class="col-3">
            <q-input
            filled
            v-model="cut_out_loss_insert"
            label="Cut out Loss"
            label-color="green-9"
            :input-style="{ color: 'black' }"
            input-class="text-center"
            dense
         
          />
          </div>
         
          <div class="col-1"></div>
           <div class="col-3">
            <q-input
            filled
            v-model="remnants_loss_insert"
            label="Remnants Loss"
            label-color="green-9"
            :input-style="{ color: 'black' }"
            input-class="text-center"
            dense
         
          />
          </div>
           </div>
           
          <br>
           <div class="row q-ma-md justify-center">
          <div class="col-1"></div>
           <div class="col-3">
            <q-input
            filled
            v-model="total_spread_loss_insert"
            label="Total Spread Loss"
            label-color="green-9"
            :input-style="{ color: 'black' }"
            input-class="text-center"
            dense
         
          />
          </div>
          <div class="col-1"></div>
           <div class="col-3">
            <q-input
            filled
            v-model="total_width_loss_insert"
            label="Total Width Loss"
            label-color="green-9"
            :input-style="{ color: 'black' }"
            input-class="text-center"
            dense
         
          />
          </div>
          </div>
         
      
        </div>
      </q-carousel-slide>
      
     
    </q-carousel>



    <div class="row justify-center">

      <q-btn v-show="this.slide == 'style'" glossy color="green-9"  label="Update" @click="Update_setting()" />

      <q-btn v-show="this.slide == 'tv'" glossy color="orange-9"  label="Add Data" @click="Insert_setting()" />
     
    </div>
    <br>

    <div class="row justify-center">
      <q-btn-toggle
        glossy
        color="blue-11"
        v-model="slide"
        :options="[
          { label: 'View/Update', value: 'style' },
          { label: 'Insert Page', value: 'tv' },
        
        ]"
      />
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
  setup () {
    return {
      slide: ref('style'),
     
    }
  },
  data(){
      return{
          type:"Year",
          end_loss : "End Loss",
          splice_loss:"Splice Loss",
          cut_out_loss:"Cut out Loss",
          remnants_loss : "Remnants Loss",
          total_spread_loss:"Total Spread Loss",
          total_width_loss:"Total Width Loss",
          year:[],
          end_loss_value:[],
          splice_loss_value:[],
          cut_out_loss_value:[],
          remnants_loss_value:[],
          total_spread_loss_per:[],
          total_width_loss_per:[],
          index_no:[],

          year_col:"Year",
          end_loss_col:"End Loss",
          splice_loss_col:"Splice Loss",
          cut_out_loss_col:"Cut out Loss",
          remnants_loss_col:"Remnants Loss",
          total_spread_loss_col:"Total spread Loss",
          total_width_loss_col:"Total width Loss",
          year_insert:"",
          end_loss_insert:"",
          splice_loss_insert:"",
          cut_out_loss_insert:"",
          remnants_loss_insert:"",
          total_spread_loss_insert:"",
          total_width_loss_insert:"",
          max_value:"",

      }
  },
  mounted() {
    let username = this.$q.localStorage.getItem("username");
    this.org = this.$q.localStorage.getItem("org");
    let status = this.$q.localStorage.getItem("status");

    this.year=[]

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
                this.index_no[i] = e.INDEX_NO
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
        params.append("index_no",this.index_no[ax])
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
    async Insert_setting(){
      
      if(this.year_insert == ''){
        this.$q.notify({
          message: "Please Insert Year" ,
          position:"center",
          color: "red-9",
        });
      }else{
      const params = new FormData(); //check year and index max
      params.append("year",this.year_insert)
      params.append("org",this.org)
       await   axios({
        method:'post',
        url: this.$api_url + '/mainso.php/check_year_duplicate',
        data:params,
      }).then((resp)=>{
          if(resp.data.data.length > 0){
          this.$q.notify({
          message: "This year has been already Please insert another value" ,
          position:"center",
          color: "red-9",
        });
         }else{
         axios({
        method:'post',
        url: this.$api_url + '/mainso.php/check_max_value',
        data:params,
      }).then((resp)=>{
        
        this.max_value = resp.data.data[0].LAST_INDEX
        console.log(this.max_value)

          const params2 = new FormData();
           params2.append("year",this.year_insert)
           params2.append("org",this.org)
           params2.append("end_loss",this.end_loss_insert)
           params2.append("splice_loss",this.splice_loss_insert)
           params2.append("cut_out_loss",this.cut_out_loss_insert)
           params2.append("remnants_loss",this.remnants_loss_insert)
           params2.append("total_spread_loss",this.total_spread_loss_insert)
           params2.append("total_width_loss",this.total_width_loss_insert)
           params2.append("max_value",Number(this.max_value)+1)
            for (var pair of params2.entries()) {
                  console.log(pair[0] + ", " + pair[1]);
          }
        axios({
        method:'post',
        url: this.$api_url + '/mainso.php/insert_new_value_year',
        data:params2,
      }).then((resp)=>{
        if(resp.data.status == true){
         this.$q.notify({
          message: "Insert Success" ,
          position:"center",
          color: "green-9",
        });
        }
      })

       
      })
           
           
         }
      }).catch((error)=>{
        console.log(error)
      })
    }
    }
  }
}
</script>
<style>


</style>