<template>
  <div class="q-pa-md" style="max-width: 100%">
    <div class="q-gutter-md">
      <q-card class="my-card" style="max-width: 100%">
        <div class="row">
          <div class="col-3">
            <q-input
              input-class="text-center"
              type="number"
              class="q-ma-xs"
              v-model="SO_YEAR"
              @keypress.enter.prevent="search()"
              dense
              label="Year"
              :bg-color="isDisabled ? 'blue-grey-1 ' : ''"
              :disable="isDisabled"
            />
          </div>

          <div class="col-3">
            <q-input
              input-class="text-center"
              class="q-ma-xs"
              v-model="SO_NO"
              @keypress.enter.prevent="search()"
              type="number"
              dense
              label="SO"
              :bg-color="isDisabled ? 'blue-grey-1 ' : ''"
              :disable="isDisabled"
            />
          </div>
          <div class="col-3">
            <q-select
              input-class="text-center"
              class="q-ma-xs"
              v-model="F_TYPE"
              :options="f_type"
              label="F TYPE"
             
              dense
            />
            <q-tooltip v-model="Showing2"
              >A: WOVEN - NYLON, TAFFETA.<br />
              B:LIGHT-MIDDLE WEIGHT - SINGLE JERSEY, INTERLOCK, DOUBLE KNIT,
              JACQUARD <br />
              C:FLEECE<br />
              D:SPECIAL - RIB, ELASTIC, ELASTAN, LYCRA, SPANDEX, TWO WAY,
              WAFFLE, PRE WASH, PRESHRUNK, VELOUR<br />
            </q-tooltip>
          </div>

          <div class="col-3">
            <q-btn
              input-class="text-center"
              class="q-ma-xs"
              v-if="this.isDisabled"
              icon="search"
              type="submit"
              size="md"
              :disable="isDisabled"
              stack
              color="blue-3 text-black"
            />

            <q-btn
              v-else
              class="q-ma-xs"
              icon="search"
              type="submit"
              label=""
              size="md"
              stack
              @click="search()"
              color="teal-8 text-white"
            >
              <q-tooltip
                transition-show="flip-right"
                transition-hide="flip-left"
              >
                Search
              </q-tooltip>
            </q-btn>

            <q-btn
              class="q-ma-xs"
              icon="close"
              type="submit"
              size="md"
              label=""
              stack
              color="warning text-black"
              @click="clearall()"
            >
              <q-tooltip
                transition-show="flip-right"
                transition-hide="flip-left"
              >
                Clear ALL
              </q-tooltip>
            </q-btn>
          </div>
        </div>

        <div class="row">
          <div class="col-3">
            <q-input
              input-class="text-center"
              class="q-ma-xs"
              v-model="posts.SO_NO_DOC"
              dense
              label="SO_NO_DOC"
              :bg-color="isDisabled ? 'blue-grey-1' : ''"
              :disable="isDisabled"
            />
          </div>

          <div class="col-3">
            <q-input
              input-class="text-center"
              class="q-ma-xs"
              v-model="posts.STYLE_REF"
              dense
              label="STYLE"
              :bg-color="isDisabled ? 'blue-grey-1' : ''"
              :disable="isDisabled"
            />
          </div>

          <div class="col-6">
            <q-input
              input-class="text-center"
              class="q-ma-xs"
              v-model="posts.CUST_NAME"
              dense
              label="CUSTOMER"
              :bg-color="isDisabled ? 'blue-grey-1' : ''"
              :disable="isDisabled"
            />
          </div>
          <div class="col-1"></div>
        </div>

        <div class="row">
          <div class="col-3">
            <q-input
              input-class="text-center"
              class="q-ma-xs"
              v-model="posts.QTY"
              dense
              label="QTY ORDER"
              :bg-color="isDisabled ? 'blue-grey-1' : ''"
              :disable="isDisabled"
            />
          </div>

          <div class="col-3">
            <q-input
              input-class="text-center"
              class="q-ma-xs"
              v-model="data_top_qty_cut.CUT_QTY"
              dense
              label="QTY CUT"
              :bg-color="isDisabled ? 'blue-grey-1' : ''"
              :disable="isDisabled"
            />
          </div>

          <div class="col-3">
            <q-input
              input-class="text-center"
              class="q-ma-xs"
              v-model="posts.GMT_TYPE"
              dense
              label="GMT TYPE"
              :bg-color="isDisabled ? 'blue-grey-1' : ''"
              :disable="isDisabled"
            />
          </div>
        </div>
        <div class="row">
          <div class="col-12">
            
          
            <q-select 
            v-model="fabric_type" 
            :options="row_fabric" 
            label="Fabric Type" />
         
         <!--    <q-input
              input-class="text-center"
              class="q-ma-xs"
              v-model="posts.FABRIC_TYPE"
              dense
              label="Fabric TYPE"
              :bg-color="isDisabled ? 'blue-grey-1' : ''"
              :disable="isDisabled"
            /> -->
           
          </div>
        </div>
      </q-card>

      <div class="row">
        <div class="col-1">
          <q-btn icon="add" size="md" color="green-9" @click="addRowgm" />
        </div>

        <div class="col-1"></div>
        <div class="col-1">
          <q-select
            outlined
            v-model="cpartno"
            :options="rowz"
            label="Cpart"
            dense
            :disable="isDisabled2"
            :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
            @onclick="checkgm()"
          />
        </div>
        <div class="col-1"></div>
         <div class="col-1">
         <q-input dense filled v-model="start">
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
        <div class="col-1"></div>
        
        <div class="col-1">
          <q-select
            outlined
            v-model="org"
            :options="orgs"
            dense
            label="Chose Org"
            disable
            
          />
        </div>
        <div class="col-1"></div>
        <div class="col-1">
          <q-input
            filled
            v-model="endloss_auto"
            type="number"
            dense
            label="Endloss auto (YD)"
          />
        </div>
        <div class="col-1"></div>
        <div class="col-1">
          <q-input
            filled
            v-model="endloss_auto_inch"
            type="number"
            dense
            label="Endloss auto(Inch)"
          />
        </div>
      </div>

      <div class="row">
        <div class="col-12">
          <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
            <q-table
              class="my-stickys"
              :rows="rowx"
              :columns="columns"
              hide-pagination
              table-header-style="text-align:center;"
              table-header-class="bg-blue-4 text-black"
              :rows-per-page-options="[10]"
              row-key="name"
              dense
              style="min-height: 100px"
            >
              <template v-slot:body="props">
                <q-tr table-class="bg-red-3" :props="props">
                  <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                  <!--   -->

                  <q-td key="name" :props="props">
                    <q-select
                      v-if="props.row.status == false"
                      outlined
                      v-model="props.row.gm_no"
                      bg-color="red-2"
                      :options="rowxy"
                      label="Chose GM"
                      dense
                      :clearable="category$clearable"
                      :clear-value="category$clear"
                      @update:model-value="gmselect(props.pageIndex)"
                      style="min-width: 100px"
                    />
                    <q-select
                      v-else
                      outlined
                      v-model="props.row.gm_no"
                      bg-color="green-2"
                      :options="rowxy"
                      label="Chose GM"
                      dense
                      :clearable="category$clearable"
                      :clear-value="category$clear"
                      @update:model-value="gmselect(props.pageIndex)"
                      style="min-width: 100px"
                    />
                    <!-- {{props.row.gm}}
                                                                    <select v-model="props.row.gm"  style="min-width: 100px;" @change="gmselect(props.row.gm)">
                                                                        <option v-for="item in gm" value="" :key="item.name" >--Select--</option>
                                                                        <option v-for="item in gm" value={{item.name}} :key="item.name" >{{item.name}}</option>
                                                                    </select> -->
                  </q-td>
                  <q-td key="number" :props="props">
                  
                    <q-select outlined v-model="props.row.number" :options="chose_table" label="Table" />
                  </q-td>
                  <q-td key="gm_m2" :props="props">
                    <q-input
                      input-class="text-center"
                      type="number"
                      v-model="props.row.GM_M2"
                      label="GM/M2"
                    />
                  </q-td>
                  <q-td key="gm_yd" :props="props">
                    <q-input
                      input-class="text-center"
                      type="number"
                      v-model="props.row.GM_YD"
                      label="gm_yd"
                    />
                  </q-td>
                  <q-td key="widthinch" :props="props">
                    <q-input
                      input-class="text-center"
                      type="number"
                      v-model="props.row.WIDTH_INCH"
                      label="widthinch"
                    />
                  </q-td>
                  <q-td key="lenthyds" :props="props">
                    <q-input
                      input-class="text-center"
                      type="number"
                      v-model="props.row.YD"
                      label="YARD"
                    />
                  </q-td>
                  <q-td key="lenthyd" :props="props">
                    <q-input
                      input-class="text-center"
                      type="number"
                      v-model="props.row.LENGTH_YD"
                      label="lenthyd"
                    />
                  </q-td>

                  <q-td key="lenthinch" :props="props">
                    <q-input
                      input-class="text-center"
                      type="number"
                      v-model="props.row.LENGTH_INCH"
                      label="lenthinch"
                    />
                  </q-td>
                  <q-td key="ttlmarker" :props="props">
                    <q-input
                      input-class="text-center"
                      type="number"
                      v-model="props.row.ttlmarker"
                      label="ttlmarker"
                    />
                  </q-td>
                  <q-td key="marker" :props="props">
                    <q-input
                      input-class="text-center"
                      type="number"
                      v-model="props.row.PER_U"
                      label="marker"
                    />
                  </q-td>
                  <q-td key="ply" :props="props">
                    <q-input
                      input-class="text-center"
                      type="number"
                      v-model="props.row.LAYER"
                      label="ply"
                    />
                  </q-td>

                  <q-td key="insert" :props="props">
                    <!--key = name.columne -->
                    <q-btn
                      v-if="props.row.status == false"
                      icon="note_add"
                      color="green-9"
                      @click="addDetail(props.pageIndex)"
                    />
                    <q-btn
                      v-else
                      icon="note_add"
                      color="light-blue-9"
                      @click="goDetail(props.pageIndex)"
                    />
                  </q-td>
                  <q-td key="delete" :props="props">
                    <!--key = name.columne -->
                    <q-btn
                      icon="delete"
                      color="red-9"
                      @click="deleteRow(props.pageIndex)"
                    />
                    <!--ต้องใช้ row.value json เสมอ -->
                  </q-td>
                </q-tr>
              </template>
            </q-table>
          </div>
        </div>
      </div>
      <br />
      <br />

      <q-tabs v-model="tab" dense align="justify" narrow-indicator>
        <q-tab
          name="main"
          class="bg-blue-10 text-white"
          label="MAIN"
          @click="calavg()"
          style="width: 90%"
        />

        <q-tab
          name="result"
          class="bg-blue-10 text-white"
          label="Result"
          @click="calavg()"
          style="width: 90%"
        />
      </q-tabs>

      <q-tab-panels v-model="tab" animated>
        <q-tab-panel name="main">
          <div class="row">
            <div class="col-12">
              <div
                class="q-my-sm"
                style="margin-top:px; padding-top: 0px padding-bottom:0px"
              >
                <q-table
                  :rows="rowmain"
                  :columns="comlumns_bottom_1"
                  table-header-class="bg-blue-3 text-black"
                  hide-bottom
                  dense
                  row-key="issue"
                  style="min-height: 150px"
                >
                  <template v-slot:body="props">
                    <q-tr :props="props">
                      <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                      <!--   -->

         
                      <q-td key="issue" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.issue"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="avg_width" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.avg_width"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="issue_roll" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.issue_roll"
                          :disable="true"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="yard_roll" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.yard_roll"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="gmm2" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.gmm2"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="gmyd" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.gmyd"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="width_inch" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.width_inch"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="length_ydb" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.length_ydb"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="length_yd" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.length_yd"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="length_inch" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.length_inch"
                          dense
                          label=""
                        />
                      </q-td>
                    </q-tr>
                  </template>
                </q-table>
              </div>

              <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
                <q-table
                  :rows="rowmain"
                  :columns="comlumns_bottom_2"
                  :rows-per-page-options="[10]"
                  table-header-class="bg-blue-3 text-black"
                  hide-pagination
                  row-key="cause"
                  dense
                  style="min-height: 150px"
                >
                  <template v-slot:body="props">
                    <q-tr :props="props">
                      <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                      <!--   -->

                     
                      <q-td key="obs_width" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.obs_width"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="marker" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.marker"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="ply" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.ply"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="ttl_marker" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.ttl_marker"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="endplug1" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.endplug1"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="endplug2" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.endplug2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="endplug12_yd" :props="props">
                        <q-input
                         
                          input-class="text-center"
                          type="number"
                          v-model="props.row.endplug12_yd"
                          dense
                         
                        
                          label=""
                        />
                        
                      </q-td>
                       
                      <q-td key="splicenoof" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.splicenoof"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="spliceinch" :props="props">
                        <q-input
                          v-if="this.save_status_splice == false"
                          input-class="text-center"
                          type="number"
                          bg-color="red-2"
                          v-model="props.row.spliceinch"
                          dense
                         
                          label=""
                        />
                        <q-input
                          v-else
                          input-class="text-center"
                          type="number"
                          v-model="props.row.spliceinch"
                          dense
                          bg-color="green-2"
                        
                          label=""
                        />
                      </q-td>
                       <q-td key="spliceinch" :props="props">
                       <q-btn
                        push
                          align="left"
                          color="primary"
                          v-if="this.save_status_splice == false"
                        
                          @click="opensplice()"
                          label="..."
                        />
                         <q-btn
                        push
                          align="left"
                          color="secondary"
                          v-else
                        
                          @click="check_splice()"
                          label="..."
                        />
                        </q-td>
                    </q-tr>
                  </template>
                </q-table>
              </div>
            </div>
          </div>
        </q-tab-panel>

        <q-tab-panel name="result">
          <div class="row">
            <div class="col-12">
              <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
                <q-table
                  table-header-class="bg-blue-3 text-black"
                  :rows="rowmain"
                  :columns="comlumns_bottom_3"
                  :rows-per-page-options="[10]"
                  row-key="yardroll"
                  hide-pagination
                  dense
                  style="min-height: 150px"
                >
                  <template v-slot:body="props">
                    <q-tr :props="props">
                      <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                      <!--   -->

                      <q-td key="cutoutweight" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.cutoutweight"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="remnants" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.remnants"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="widthplug1" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.widthplug1"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="widthplug2" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.widthplug2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="plug12_inch" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.plug12_inch"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="end1_2" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.end1_2"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="endloss" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.endloss"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="plug1_inch" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.plug1_inch"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="per_plug1" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.per_plug1"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="plug2_inch" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.plug2_inch"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="per_plug2" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.per_plug1"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                    </q-tr>
                  </template>
                </q-table>
              </div>
              <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
                <q-table
                  table-header-class="bg-blue-3 text-black"
                  :rows="rowmain"
                  :columns="comlumns_bottom_4"
                  :rows-per-page-options="[10]"
                  row-key="sectionlosslenth"
                  hide-pagination
                  dense
                  style="min-height: 150px"
                >
                  <template v-slot:body="props">
                    <q-tr :props="props">
                      <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                      <!--   -->

                      <q-td key="endloss_yd" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.endloss_yd"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="avgglueside" :props="props">
                        <q-input
                          v-if="this.save_status_glue_side == false"
                          bg-color="red-2"
                          input-class="text-center"
                          type="number"
                          v-model="props.row.avgglueside"
                         
                          dense
                          label=""
                        />
                        <q-input
                          v-else
                          input-class="text-center"
                          bg-color="green-2"
                          type="number"
                          v-model="props.row.avgglueside"
                         
                          dense
                          label=""
                        />
                      </q-td>

                      <q-td key="avgglueside" :props="props">
                        <q-btn
                        push
                          align="left"
                          color="primary"
                          v-if="this.save_status_glue_side == false"
                        
                          @click="add_glue_side()"
                          label="..."
                        />
                        <q-btn
                        push
                          align="left"
                          color="secondary"
                          v-else
                          
                          @click="check_glue_side()"
                          label="..."
                        />
                      </q-td>

                      <q-td key="plug12" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.plug12"
                         
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="widthloss" :props="props">
                        <q-input
                           v-if="
                            this.rowmain[0].widthloss > this.fix_width_loss
                          "
                          :input-style="{ color: 'red' }"
                          input-class="text-center"
                          type="number"
                          v-model="props.row.widthloss"
                          @click="open_widthloss()"
                          dense
                          label=""
                        />
                          <q-input
                          v-else
                          input-class="text-center"
                          type="number"
                          v-model="props.row.widthloss"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="endless_length_yd" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.endless_length_yd"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                        
                      </q-td>
                      <q-td key="avg_end" :props="props">
                        <q-input
                          v-if="this.rowmain[0].avg_end > this.fix_avg_end_loss"
                          :input-style="{ color: 'red' }"
                          input-class="text-center"
                          type="number"
                          v-model="props.row.avg_end"
                          @click="open_avg()"
                          dense
                          label=""
                        />
                        <q-input
                          v-else
                          input-class="text-center"
                          type="number"
                          v-model="props.row.avg_end"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="sectionlosslenth" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.sectionlosslenth"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="sectionlossper" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.sectionlossper"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="spliceloss" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.spliceloss"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="splicelossper" :props="props">
                        <q-input
                          v-if="
                            this.rowmain[0].splicelossper > this.fix_splice_loss
                          "
                          :input-style="{ color: 'red' }"
                          input-class="text-center"
                          type="number"
                          v-model="props.row.splicelossper"
                          @click="open_splice_loss_per()"
                          dense
                          label=""
                        />
                        <q-input
                          v-else
                          input-class="text-center"
                          type="number"
                          v-model="props.row.splicelossper"
                          dense
                          label=""
                        />
                      </q-td>
                    </q-tr>
                  </template>
                </q-table>
              </div>
              <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
                <q-table
                  table-header-class="bg-blue-3 text-black"
                  :rows="rowmain"
                  :columns="comlumns_bottom_5"
                  :rows-per-page-options="[10]"
                  row-key="defect"
                  hide-pagination
                  dense
                  style="min-height: 150px"
                >
                  <template v-slot:body="props">
                    <q-tr :props="props">
                      <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                      <!--   -->

                      <q-td key="overlap" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.overlap"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="overlapper" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.overlapper"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="defect" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.defect"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="defectper" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.defectper"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="totalcutout" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.totalcutout"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="totalcutoutper" :props="props">
                        <q-input
                          v-if="
                            this.rowmain[0].totalcutoutper >
                            this.fix_total_cut_out_per
                          "
                          :input-style="{ color: 'red' }"
                          input-class="text-center"
                          type="number"
                          v-model="props.row.totalcutoutper"
                          @click="open_total_cut_out_per()"
                          dense
                          label=""
                        />
                        <q-input
                          v-else
                          input-class="text-center"
                          type="number"
                          v-model="props.row.totalcutoutper"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="rement" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.rement"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>

                      <q-td key="rement_per" :props="props">
                        <q-input
                          v-if="
                            this.rowmain[0].rement_per > this.fix_remnants_loss
                          "
                          :input-style="{ color: 'red' }"
                          input-class="text-center"
                          type="number"
                          v-model="props.row.rement_per"
                          @click="open_remnants_loss()"
                          dense
                          label=""
                        />
                        <q-input
                          v-else
                          input-class="text-center"
                          type="number"
                          v-model="props.row.rement_per"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="totallenthloss" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.totallenthloss"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="totallenthlossper" :props="props">
                        <q-input
                          v-if="
                            this.rowmain[0].totallenthlossper > this.fix_total_length_loss
                          "
                          :input-style="{ color: 'red' }"
                          input-class="text-center"
                          type="number"
                          v-model="props.row.totallenthlossper"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                        <q-input
                          v-else
                          input-class="text-center"
                          type="number"
                          v-model="props.row.totallenthlossper"
                          :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                          :disable="isDisabled2"
                          dense
                          label=""
                        />
                      </q-td>
                      <q-td key="re_cut" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.re_cut"
                          dense
                          label=""
                        />
                      </q-td>
                    </q-tr>
                  </template>
                </q-table>
              </div>
            </div>
          </div>
        </q-tab-panel>
      </q-tab-panels>

      <!--  <q-form @submit="Savedata()"> -->

      <q-dialog
        v-model="basic"
        transition-show="rotate"
        transition-hide="rotate"
      >
        <q-card style="min-width: 50%">
          <q-card-section>
            <div class="text-h6">Insert ข้อมูล</div>
            <q-card-actions align="right">
              <q-btn icon="add" color="positive" size="md" @click="addRow" />
              <q-btn
                v-if="this.save_status == false"
                icon="save"
                size="md"
                color="teal-10"
                type="submit"
                @click="Savedata()"
                v-close-popup
              >
              </q-btn>
              <q-btn
                v-else
                icon="save"
                size="md"
                color="warning"
                type="submit"
                @click="update_detail_sub()"
                v-close-popup
              >
              </q-btn>
            </q-card-actions>
          </q-card-section>

          <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
            <q-table
              :rows="rowwidth"
              :columns="columns3"
              class="my-sticky"
              table-header-class="bg-light-blue-3"
              :rows-per-page-options="[20]"
              hide-pagination=""
              row-key="no"
              dense
              style="height: 90%"
            >
              <template v-slot:body="props">
                <q-tr :props="props">
                  <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                  <!--   -->

                  <q-td key="no" :props="props">
                    {{ props.pageIndex + 1 }}
                  </q-td>
                   <q-td key="issue_roll" :props="props">
                    <q-input v-show="props.pageIndex <= 0"
                      input-class="text-center"
                      type="number"
                      v-model="props.row.issue_roll"
                      dense
                      label=""
                    />
                    
                  </q-td>
                  <q-td key="issue_yard" :props="props">
                   <q-input v-show="props.pageIndex <= 0"
                      input-class="text-center"
                      type="number"
                      v-model="props.row.issue_yard"
                      @update:model-value="calculate_issue(props.pageIndex)"
                      dense
                      label=""
                    /> 

                  
                  </q-td>
                 
                  <q-td key="issue_kg" :props="props">
                     <q-input v-show="props.pageIndex <= 0"
                      input-class="text-center"
                      type="number"
                      v-model="props.row.issue_kg"
                      disable
                      dense
                      label=""
                    />
                  </q-td>
                  <q-td key="avg_width" :props="props">
                    <q-input
                      input-class="text-center"
                      type="number"
                      v-model="props.row.avg_width"
                      dense
                      label=""
                    />
                  </q-td>
                  <q-td key="delete" :props="props">
                    <!--key = name.columne -->
                    <q-btn
                      color="red-12"
                      icon="delete"
                      size="md"
                      @click="deleteSelected(props.pageIndex)"
                    />
                  </q-td>
                </q-tr>
              </template>
            </q-table>
          </div>
        </q-card>
      </q-dialog>

      <q-dialog
        v-model="status_glue"
        transition-show="rotate"
        transition-hide="rotate"
      >
        <q-card style="min-width: 50%">
          <q-card-section>
            <div class="text-h6">Insert ข้อมูล Avg Glue Side</div>
            <q-card-actions align="right">
              <q-btn
                v-if="save_status_glue_side == false"
                icon="save"
                size="md"
                color="teal-10"
                type="submit"
                @click="Save_data_glue_side()"
                v-close-popup
              />
              <q-btn
                v-else
                icon="save"
                size="md"
                color="warning"
                type="submit"
                @click="Update_data_glue_side()"
                v-close-popup
              >
              </q-btn>
            </q-card-actions>
          </q-card-section>

          <div
            class="q-ma-none q-pa-none"
            id="glue_side"
            style="margin-top: 0px; padding-top: 0px"
          >
            <q-table
              :rows="row_glue_side"
              :columns="columns_glue_side"
              class="my-sticky"
              table-header-class="bg-light-blue-3"
              :rows-per-page-options="[20]"
              hide-pagination=""
              row-key="no"
              dense
              style="height: 380px"
            >
              <template v-slot:body="props">
                <q-tr :props="props">
                  <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                  <!--   -->

                  <q-td key="no" :props="props">
                    {{ props.pageIndex + 1 }}
                  </q-td>

                  <q-td key="glue_side1" :props="props">
                    <q-input
                      input-class="text-center"
                      type="number"
                      v-model="props.row.glue_side1"
                      dense
                      label=""
                    />
                  </q-td>
                  <q-td key="glue_side2" :props="props">
                    <q-input
                      input-class="text-center"
                      type="number"
                      v-model="props.row.glue_side2"
                      dense
                      label=""
                    />
                  </q-td>
                </q-tr>
              </template>
            </q-table>
          </div>
        </q-card>
      </q-dialog>

      <q-dialog
        v-model="status_splice"
        transition-show="rotate"
        transition-hide="rotate"
      >
        <q-card style="min-width: 20%">
          <q-card-section>
            <div class="text-h6">Insert ข้อมูล Splice Inch</div>
            <q-card-actions align="right">
              <q-btn
                icon="add"
                v-if="save_status_splice == true"
                color="positive"
                size="md"
                @click="addRowsplice()"
              />
              <q-btn
                v-if="save_status_splice == false"
                icon="save"
                size="md"
                color="teal-10"
                type="submit"
                @click="Save_data_splice()"
                v-close-popup
              >
              </q-btn>
              <q-btn
                v-else
                icon="save"
                size="md"
                color="warning"
                type="submit"
                @click="Update_data_splice()"
                v-close-popup
              >
              </q-btn>
            </q-card-actions>
          </q-card-section>

          <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
            <q-table
              :rows="row_splice"
              :columns="columns_splice"
              class="my-sticky"
              table-header-class="bg-light-blue-3"
              :rows-per-page-options="[20]"
              hide-pagination=""
              row-key="no"
              dense
              style="height: 380px"
            >
              <template v-slot:body="props">
                <q-tr :props="props">
                  <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                  <!--   -->

                  <q-td key="no" :props="props">
                    {{ props.pageIndex + 1 }}
                  </q-td>

                  <q-td key="splice_value" :props="props">
                    <q-input
                      input-class="text-center"
                      type="number"
                      v-model="props.row.splice_value"
                      dense
                      label=""
                    />
                  </q-td>
                </q-tr>
              </template>
            </q-table>
          </div>
        </q-card>
      </q-dialog>
      <q-dialog
        v-model="status_avg_end"
        transition-show="rotate"
        transition-hide="rotate"
      >
        <q-card style="min-width: 60%">
          <q-card-section>
            <div class="text-h6">Insert ข้อมูล AVG END</div>
            <q-card-actions align="right">
              <q-btn
                icon="save"
                size="md"
                color="teal-10"
                type="submit"
                @click="Save_data_avg_end()"
                v-close-popup
              >
              </q-btn>
            </q-card-actions>
          </q-card-section>

          <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
            <q-table
              :rows="row_avg_end"
              :columns="columns_avg_end"
              class="my-sticky"
              table-header-class="bg-light-blue-3"
              :rows-per-page-options="[20]"
              hide-pagination=""
              row-key="no"
              dense
              style="height: 380px; min-width: 700px"
            >
              <template v-slot:body="props">
                <q-tr :props="props">
                  <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                  <!--   -->

                  <q-td key="avg_end_code" :props="props">
                    <q-select
                      filled
                      v-model="avg_end_code"
                      :options="option_avg_end_code"
                      label=""
                    />
                  </q-td>
                </q-tr>
              </template>
            </q-table>
          </div>
        </q-card>
      </q-dialog>
      <q-dialog
        v-model="status_splice_loss_per"
        transition-show="rotate"
        transition-hide="rotate"
      >
        <q-card style="min-width: 60%">
          <q-card-section>
            <div class="text-h6">Insert ข้อมูล SPLICE LOSS</div>
            <q-card-actions align="right">
              <q-btn
                icon="save"
                size="md"
                color="teal-10"
                type="submit"
                @click="Save_splice_loss_per()"
                v-close-popup
              >
              </q-btn>
            </q-card-actions>
          </q-card-section>

          <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
            <q-table
              :rows="row_per_splice_loss"
              :columns="columns_per_splice_loss"
              class="my-sticky"
              table-header-class="bg-light-blue-3"
              :rows-per-page-options="[20]"
              hide-pagination=""
              row-key="no"
              dense
              style="height: 380px; min-width: 700px"
            >
              <template v-slot:body="props">
                <q-tr :props="props">
                  <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                  <!--   -->

                  <q-td key="per_splice_loss_code" :props="props">
                    <q-select
                      filled
                      v-model="per_splice_loss_code"
                      :options="option_per_splice_loss"
                      label=""
                    />
                  </q-td>
                </q-tr>
              </template>
            </q-table>
          </div>
        </q-card>
      </q-dialog>

      <q-dialog
        v-model="status_total_cut_out_per"
        transition-show="rotate"
        transition-hide="rotate"
      >
        <q-card style="min-width: 60%">
          <q-card-section>
            <div class="text-h6">Insert ข้อมูล TOTAL CUT OUT PER</div>
            <q-card-actions align="right">
              <q-btn
                icon="save"
                size="md"
                color="teal-10"
                type="submit"
                @click="Save_total_cut_out_per()"
                v-close-popup
              >
              </q-btn>
            </q-card-actions>
          </q-card-section>

          <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
            <q-table
              :rows="row_total_cut_out_per"
              :columns="columns_total_cut_out_per"
              class="my-sticky"
              table-header-class="bg-light-blue-3"
              :rows-per-page-options="[20]"
              hide-pagination=""
              row-key="no"
              dense
              style="height: 380px; min-width: 700px"
            >
              <template v-slot:body="props">
                <q-tr :props="props">
                  <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                  <!--   -->

                  <q-td key="total_cut_out_per_code" :props="props">
                    <q-select
                      filled
                      v-model="total_cut_out_per_code"
                      :options="option_total_cut_out_per"
                      label=""
                    />
                  </q-td>
                </q-tr>
              </template>
            </q-table>
          </div>
        </q-card>
      </q-dialog>

      <q-dialog
        v-model="status_remnants_loss"
        transition-show="rotate"
        transition-hide="rotate"
      >
        <q-card style="min-width: 60%">
          <q-card-section>
            <div class="text-h6">Insert ข้อมูล REMNANTS LOSS</div>
            <q-card-actions align="right">
              <q-btn
                icon="save"
                size="md"
                color="teal-10"
                type="submit"
                @click="Save_remnants_loss()"
                v-close-popup
              >
              </q-btn>
            </q-card-actions>
          </q-card-section>

          <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
            <q-table
              :rows="row_remnants_loss"
              :columns="columns_remnants_loss"
              class="my-sticky"
              table-header-class="bg-light-blue-3"
              :rows-per-page-options="[20]"
              hide-pagination=""
              row-key="no"
              dense
              style="height: 380px; min-width: 700px"
            >
              <template v-slot:body="props">
                <q-tr :props="props">
                  <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                  <!--   -->

                  <q-td key="remnants_loss_code" :props="props">
                    <q-select
                      filled
                      v-model="remnants_loss_code"
                      :options="option_remnants_loss"
                      label=""
                    />
                  </q-td>
                </q-tr>
              </template>
            </q-table>
          </div>
        </q-card>
      </q-dialog>

      <q-dialog
        v-model="status_widthloss"
        transition-show="rotate"
        transition-hide="rotate"
      >
        <q-card style="min-width: 60%">
          <q-card-section>
            <div class="text-h6">Insert ข้อมูล WIDTH LOSS</div>
            <q-card-actions align="right">
              <q-btn
                icon="save"
                size="md"
                color="teal-10"
                type="submit"
                @click="Save_widthloss()"
                v-close-popup
              >
              </q-btn>
            </q-card-actions>
          </q-card-section>

          <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
            <q-table
              :rows="row_widthloss"
              :columns="columns_widthloss"
              class="my-sticky"
              table-header-class="bg-light-blue-3"
              :rows-per-page-options="[20]"
              hide-pagination=""
              row-key="no"
              dense
              style="height: 380px; min-width: 700px"
            >
              <template v-slot:body="props">
                <q-tr :props="props">
                  <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                  <!--   -->

                  <q-td key="widthloss_code" :props="props">
                    <q-select
                      filled
                      v-model="widthloss_code"
                      :options="option_widthloss"
                      label=""
                    />
                  </q-td>
                </q-tr>
              </template>
            </q-table>
          </div>
        </q-card>
      </q-dialog>
      <div class="q-pa-md q-gutter-sm">
        <q-dialog
          v-model="confirm_delete_detail"
          transition-show="rotate"
          transition-hide="rotate"
        >
          <q-card style="min-width: 30%">
            <q-card-section>
              <div class="text-h6">Confirm Delete</div>
              <q-card-actions align="right">
                <q-btn color="green-4" size="lg" icon="done" @click="confirm_delete()"/>
                <q-btn color="negative" size="lg" icon="cancel"  @click="cancel_delete()" />
              </q-card-actions>
            </q-card-section>

           
          </q-card>
        </q-dialog>
      </div>

      <q-card-actions align="center" class="q-ma-none q-pa-none">
        <q-btn
          color="indigo-10"
          label="Save Data"
          size="20px"
          @click="saveAll()"
        />
      </q-card-actions>
    </div>
  </div>
</template>

<script>
import { ref } from "vue";
import axios from "axios";
import { colors } from "quasar";
import { useQuasar } from "quasar";
import moment from 'moment';

const { getPaletteColor } = colors;
const $q = useQuasar();

export default {
  setup() {
    return {
      model: ref(null),
      tab: ref("main"),
      model2: ref(null),
      status_add_row:true,

      org: ref(""),
      option_avg_end_code: [
        "EA ตั้งเผื่อเกินมาตรฐาน",
        "EB เผื่อ End เกินมาตรฐาน",
        "EC วางผ้าไม่สม่ำเสมอ",
        "ED ปัญหาผ้า เช่น ผ้าม้วน ผ้าริมตึงกลางหย่อน ตรงกลางโค้ง",
        "EE ผิดพลาดจากพนักงาน เช่น ไม่ใช้อุปกรณ์ช่วย, พนักงานเข้าใจผิด, ใช้มาร์คผิด",
        "EF ผิดพลาดจากอุปกรณ์ เครื่องจักร เช่น มาร์คสั้นหรือยาวกว่าความจริง เครื่องปูผ้าตั้งความยาวไม่ตรงความจริง",
        "EG ใช้มาร์คสั้นกว่ามาตราฐานที่กำหนด",
      ],
      option_per_splice_loss: [
        "SA  ทำการต่อเกยจำนวนมาก",
        "SB วาดรอยต่อเกยเกินจากระยะจริง",
        "SC พนักงานปูทำการต่อเกยเกินจากระยะที่กำหนด",
        "SD พนักงานปูกำหนดรอยต่อเกยเอง",
      ],
      option_total_cut_out_per: [
        "CA ตัดผ้าตำหนิจำนวนมาก",
        "CB พนักงานตัดผ้าตำหนิไม่เป็นไปตามมาตรฐานที่กำหนด",
      ],
      option_remnants_loss: [
        "RA พนักงานตัดหัว/ท้ายผ้าไม่เป็นไปตามมาตรฐานที่กำหนด",
        "RB ปัญหาผ้า เช่น มีตำหนิกว้างมาก หัวผ้าเอียง",
      ],
      option_widthloss: [
        "A แก้ไขมาร์คไม่มีผล",
        "B ไม่ได้คัดแยกหน้าผ้าปู หน้าผ้าไม่สม่ำเสมอต่างพับ",
        "C หน้าผ้าไม่สม่ำเสมอในพับเดียวกัน",
        "D ใช้มาร์คไม่เหมาะสมกับหน้าผ้า",
        "E ปัญหาผ้า เช่น ริมเสียบางช่วง, ริมย้วย, ริมม้วน, รูเข็มไม่เท่ากัน",
        "F ปูผ้าหลายสีหน้าผ้าไม่เท่ากัน",
      ],
      options: ["xx-xxxx", "xx-xxxx", "xx-xxxx", "xx-xxxx", "xxx-xxx"],
      optionsx: ["111-585758", "xx-xxxx", "xx-xxxx", "xx-xxxx", "9867-28579"],
      code: [
        "-",
        "A มาร์คไม่มีผล",
        "B ไม่ได้คัดแยกหน้าผ้าปู",
        "C หน้าผ้าไม่สม่ำเสมอในม้วน และต่างม้วน",
        "D ใช้มาร์คไม่เหมาะสมกับหน้าผ้า",
        "E อื่น ๆ ริมเสียบางช่วง, ริมย้วย, ปูผ้าหลายสีหน้าผ้าไม่เท่ากัน",
      ],
      orgs: ["G1", "G2", "G3", "G4"],
      chose_table: ["A1", "A2", "A3", "A4","A5","A6","A7","A8","A9","A10","A11","A12",
      "A13","A14","A15","A16","A17","A18","A19","A20","M1","M2","M3","M4","M5","M6","M7","M8","M9","M10"],
      f_type: [
        "1. WOVEN - NYLON, TAFFETA",
        "2. LIGHT-MIDDLE WEIGHT - SINGLE JERSEY, INTERLOCK, DOUBLE KNIT, JACQUARD",
        "3. FLEECE",
        "4. SPECIAL - RIB, ELASTIC, ELASTAN, LYCRA, SPANDEX, TWO WAY, WAFFLE, PRE WASH, PRESHRUNK, VELOUR",
      ],

      teal: ref(true),
      orange: ref(false),
      red: ref(false),
      cyan: ref(true),
      pink: ref(true),
      basic: ref(false),
      status_glue: ref(false),
      basic2: ref(false),
      text: ref(""),
      SO_YEAR: ref(null),
      SO_NO: ref(null),
      showing: ref(false),
      showing2: ref(false),
      endloss_auto: ref(""),
      endloss_auto_inch: ref(""),
      confirm_delete_detail:ref(false),
      count_row_gm:1,

      cpart: ["1", "2"],
      gm: [
        "100010101",
        "100200101",
        "210202101",
        "100202020",
        "10202490",
        "102029102",
      ],
      gmtype: ["BOTTB", "BOTTC", "OVS", "JKT"],
      ftype: [
        "1  WOVEN - NYLON, TAFFETA",
        "2 LIGHT-MIDDLE WEIGHT - SINGLE JERSEY, INTERLOCK, DOUBLE KNIT, JACQUARD ",
        "3 FLEECE",
        "4 SPECIAL - RIB, ELASTIC, ELASTAN, LYCRA, SPANDEX, TWO WAY, WAFFLE, PRE WASH, PRESHRUNK, VELOUR",
      ],
    };
  },
  data() {
    return {
      status: false,
      status2: false,
      check_status_wrong:ref(true),
      fabric_type:"",
      status_avg_end: false,
      avg_end_code: ref(""),
      avg_end_cause: ref(""),
      status_splice_loss_per: false,
      per_splice_loss_code: ref(""),
      per_splice_loss_cause: ref(""),
      status_total_cut_out_per: false,
      total_cut_out_per_code: ref(""),
      total_cut_out_per_cause: ref(""),
      status_remnants_loss: false,
      remnants_loss_code: ref(""),
      remnants_loss_cause: ref(""),
      fix_width_loss:1.50,
      fix_avg_end_loss: 0.34,
      fix_splice_loss: 0.10,
      fix_total_cut_out_per: 0.00,
      fix_remnants_loss: 0.08,
      fix_total_length_loss:0.52,
      status_widthloss: false,
      widthloss_code: ref(""),
      widthloss_cause: ref(""),
      status_check_gm: false,
      delete_index:'',
      keep_layer:[],
      count_layer:[],
      row_fabric:[],
      rowexport:[],
      start:new Date(),
      isDisabled3:false,

      columns: [
        {
          name: "name",
          label: "GM08",
          align: "center",
          field: "name",
        },
        {
          name: "number",
          label: "Table",
          align: "center",
          field: "number",
        },
        {
          name: "gm_m2",
          label: "GM_M2",
          align: "center",
          field: "gm_m2",
        },
        {
          name: "gm_yd",
          label: "GM_YD",
          align: "center",
          field: "gm_yd",
        },
        {
          name: "widthinch",
          label: "Weight Inch",
          align: "center",
          field: "weightyd",
        },

        {
          name: "lenthyds",
          label: "YARD",
          align: "center",
          field: "lenthyds",
        },
        {
          name: "lenthyd",
          label: "Lenth YD",
          align: "center",
          field: "lenthyd",
        },
        {
          name: "lenthinch",
          label: "Lenth Inch",
          align: "center",
          field: "lenthinch",
        },
        {
          name: "ttlmarker",
          label: "TTLMarker",
          align: "center",
          field: "ttlmarker",
        },
        {
          name: "marker",
          label: "Marker",
          align: "center",
          field: "marker",
        },
        {
          name: "ply",
          label: "PLY",
          align: "center",
          field: "ply",
        },
        {
          name: "insert",
          label: "Insert",
          align: "center",
          field: "insert",
        },
        {
          name: "delete",
          label: "Delete",
          align: "center",
          field: "delete",
        },
      ],
      rowx: [
        {
          name: "",
          number: "",
          weightm2: "",
          weightyd: "",
          widthinch: "",
          lenthyd: "",
          lenthyds: "",
          lenthinch: "",
          ttlmarker: "",
          marker: "",
          ply: "",
          issue: "",
          avgwidth: "",
          model: "",
          status: false,
          status2: false,
        },
      ],
      comlumns_bottom_1: [
       
        {
          name: "issue",
          label: "Issue",
          align: "center",
          field: "issue",
        },
        {
          name: "avg_width",
          label: "Avg_width",
          align: "center",
          field: "avg_width",
        },
        {
          name: "issue_roll",
          label: "Issue Roll",
          align: "center",
          field: "issue_roll",
        },
        {
          name: "yard_roll",
          label: "Issue-Yard/Roll",
          align: "center",
          field: "yard_roll",
        },

        {
          name: "gmm2",
          label: "GM/M2",
          align: "center",
          field: "gmm2",
        },
        {
          name: "gmyd",
          label: "GM/YD",
          align: "center",
          field: "gmyd",
        },
        {
          name: "width_inch",
          label: "Width Inch",
          align: "center",
          field: "width_inch",
        },
        {
          name: "length_ydb",
          label: "Yard",
          align: "center",
          field: "length_ydb",
        },
          {
          name: "length_yd",
          label: "Length yd",
          align: "center",
          field: "length_yd",
        },
        {
          name: "length_inch",
          label: "Length.Inch",
          align: "center",
          field: "length_inch",
        },
        
      ],

      comlumns_bottom_2: [
      
        {
          name: "obs_width",
          label: "Obs.Width",
          align: "center",
          field: "obs_width",
        },
        {
          name: "marker",
          label: "Marker",
          align: "center",
          field: "marker",
        },
        {
          name: "ply",
          label: "Ply",
          align: "center",
          field: "ply",
        },
        {
          name: "ttl_marker",
          label: "TTL_MARKER",
          align: "center",
          field: "ttl_marker",
        },
        {
          name: "endplug1",
          label: "End Plug1(Gram)",
          align: "center",
          field: "endplug1",
        },
        {
          name: "endplug2",
          label: "End Plug2(Gram)",
          align: "center",
          field: "endplug2",
        },
        {
          name: "endplug12_yd",
          label: "END PLUG 1+2 (YD)",
          align: "center",
          field: "endplug12_yd",
        },
       
        {
          name: "splicenoof",
          label: "Splice No of",
          align: "center",
          field: "splicenoof",
        },
        {
          name: "spliceinch",
          label: "Splice Inch",
          align: "center",
          field: "spliceinch",
        },
        {
          name: "spliceinch",
          label: "Detail Splice inch",
          align: "center",
          field: "spliceinch",
        },
      ],
      comlumns_bottom_3: [
        {
          name: "cutoutweight",
          label: "Cut out Weight",
          align: "center",
          field: "cutoutweight",
        },
        {
          name: "remnants",
          label: "Remnants",
          align: "center",
          field: "remnants",
        },
        {
          name: "widthplug1",
          label: "Width Plug1",
          align: "center",
          field: "widthplug1",
        },
        {
          name: "widthplug2",
          label: "Width Plug2",
          align: "center",
          field: "widthplug2",
        },
        {
          name: "plug12_inch",
          label: "PLUG1+2(INCH)",
          align: "center",
          field: "plug12_inch",
        },
        {
          name: "end1_2",
          label: "End1_2",
          align: "center",
          field: "end1_2",
        },
        {
          name: "endloss",
          label: "Endloss",
          align: "center",
          field: "endloss",
        },
        {
          name: "plug1_inch",
          label: "Plug1 inch",
          align: "center",
          field: "plug1_inch",
        },
        {
          name: "per_plug1",
          label: "% of Plug1",
          align: "center",
          field: "per_plug1",
        },
        {
          name: "plug2_inch",
          label: "Plug2 inch",
          align: "center",
          field: "plug2_inch",
        },
        {
          name: "per_plug2",
          label: "% of Plug2",
          align: "center",
          field: "per_plug2",
        },
      ],
      comlumns_bottom_4: [
        {
          name: "endloss_yd",
          label: "Endloss_YD",
          align: "center",
          field: "endloss_yd",
        },
        {
          name: "avgglueside",
          label: "Avg.Glue Side",
          align: "center",
          field: "avgglueside",
        },
        {
          name: "detail a.glue side",
          label: "Detail a.glue side",
          align: "center",
          field: "avgglueside",
        },
        {
          name: "plug12",
          label: "Plug1+2 No Width",
          align: "center",
          field: "Plug12",
        },
        {
          name: "widthloss",
          label: "Width Loss %",
          align: "center",
          field: "widthloss",
        },
        {
          name: "endless_length_yd",
          label: "Endloss length YD",
          align: "center",
          field: "endless_length_yd",
        },

        {
          name: "avg_end",
          label: "End Loss %",
          align: "center",
          field: "avg_end",
        },

        {
          name: "sectionlosslenth",
          label: "Section Loss Length",
          align: "center",
          field: "sectionlosslenth",
        },
        {
          name: "sectionlossper",
          label: "Section Loss %",
          align: "center",
          field: "sectionlossper",
        },
        {
          name: "spliceloss",
          label: "Splice Loss",
          align: "center",
          field: "spliceloss",
        },
        {
          name: "splicelossper",
          label: "SPLICE_LOSS%",
          align: "center",
          field: "splicelossper",
        },
      ],
      comlumns_bottom_5: [
        {
          name: "overlap",
          label: "Overlap",
          align: "center",
          field: "overlap",
        },
        {
          name: "overlapper",
          label: "%of OVERLAP",
          align: "center",
          field: "overlapper",
        },
        {
          name: "defect",
          label: "Defect",
          align: "center",
          field: "defect",
        },
        {
          name: "defectper",
          label: "% of defect",
          align: "center",
          field: "defectper",
        },
        {
          name: "totalcutout",
          label: "Cut Out Loss",
          align: "center",
          field: "totalcutout",
        },
        {
          name: "totalcutoutper",
          label: "Cut Out Loss %",
          align: "center",
          field: "totalcutoutper",
        },
        {
          name: "rement",
          label: "Remenants Loss Lenth",
          align: "center",
          field: "rement",
        },
        {
          name: "rement_per",
          label: "Remnants Loss %",
          align: "center",
          field: "rement_per",
        },
        {
          name: "totallenthloss",
          label: "Total Lenth Loss",
          align: "center",
          field: "totallenthloss",
        },
        {
          name: "totallenthlossper",
          label: "Total Lenth Loss %",
          align: "center",
          field: "totallenthlossper",
        },
        {
          name: "re_cut",
          label: "RE CUT",
          align: "center",
          field: "re_cut",
        },
      ],

      columns3: [
        {
          name: "no",
          label: "No.",
          align: "center",
          field: "no",
        },

        {
          name: "issue_roll",
          label: "issue_roll",
          align: "center",
          field: "issue_roll",
        },
        {
          name: "issue_yard",
          label: "Issue-Yard",
          align: "center",
          field: "issue_yard",
        },
        {
          name: "issue_kg",
          label: "Issue-KG",
          align: "center",
          field: "issue_kg",
        },
        {
          name: "avg_width",
          label: "Avg_width",
          align: "center",
          field: "avg_width",
        },
        {
          name: "delete",
          label: "Delete",
          align: "center",
          field: "delete",
        },
      ],
      rowwidth: [
        {
          no: "1",
          issue_yard: "",
          issue_kg: "",
          avg_width: "",
          delete: "",
        },

        {
          no: "2",
          issue_yard: "",
          issue_kg: "",
          avg_width: "",
          delete: "",
        },
        {
          no: "3",
          issue_yard: "",
          issue_kg: "",
          avg_width: "",
          delete: "",
        },
        {
          no: "4",
          issue_yard: "",
          issue_kg: "",
          avg_width: "",
          delete: "",
        },
        {
          no: "5",
          issue_yard: "",
          issue_kg: "",
          avg_width: "",
          delete: "",
        },
      ],
      columns_glue_side: [
        {
          name: "no",
          label: "No.",
          align: "center",
          field: "no",
        },
        {
          name: "glue_side1",
          label: "Glue side1",
          align: "center",
          field: "glue_side1",
        },
        {
          name: "glue_side2",
          label: "Glue side2",
          align: "center",
          field: "glue_side2",
        },
      ],
      row_glue_side: [
        {
          no: "1",
          glue_side1: "",
          glue_side2: "",
          delete: "",
        },
        {
          no: "2",
          glue_side1: "",
          glue_side2: "",
          delete: "",
        },
        {
          no: "3",
          glue_side1: "",
          glue_side2: "",
          delete: "",
        },
      ],
      columns_splice: [
        {
          name: "no",
          label: "No.",
          align: "center",
          field: "no",
        },
        {
          name: "splice_value",
          label: "Splice Inch",
          align: "center",
          field: "splice_value",
        },
      ],
      row_splice: [
        {
          no: "1",
          splice_value: "",
        },
        {
          no: "2",
          splice_value: "",
        },
        {
          no: "3",
          splice_value: "",
        },
        {
          no: "4",
          splice_value: "",
        },
      ],
      columns_avg_end: [
        {
          name: "avg_end_code",
          label: "AVG END CODE",
          align: "center",
          field: "avg_end_code",
        },
      ],
      row_avg_end: [
        {
          avg_end_code: "",
        },
      ],
      columns_per_splice_loss: [
        {
          name: "per_splice_loss_code",
          label: "SPLICE_LOSS_PER",
          align: "center",
          field: "per_splice_loss_code",
        },
      ],
      row_per_splice_loss: [
        {
          per_splice_loss_code: "",
        },
      ],
      columns_total_cut_out_per: [
        {
          name: "total_cut_out_per_code",
          label: "TOTAL_CUT_OUT_PER",
          align: "center",
          field: "total_cut_out_per_code",
        },
      ],
      row_total_cut_out_per: [
        {
          total_cut_out_per_code: "",
        },
      ],
      columns_remnants_loss: [
        {
          name: "remnants_loss_code",
          label: "REMNANTS LOSS",
          align: "center",
          field: "remnants_loss_code",
        },
      ],
      row_remnants_loss: [
        {
          remnants_loss_code: "",
        },
      ],
      columns_widthloss: [
        {
          name: "widthloss_code",
          label: "WIDTH LOSS",
          align: "center",
          field: "widthloss_code",
        },
      ],
      row_widthloss: [
        {
          widthloss_code: "",
        },
      ],

      rowy: [
        {
          endloss: "",
          widthloss: "",
          avglenth: "",
          avglenthper: "",
          avgend: "",
          sectionlosslenth: "",
          spliceloss: "",
          splicelossper: "",
          overlap: "",
          overlapper: "",
        },
      ],
      rowb: [
        {
          yardroll: "77.13",
          lengthyd: "5.37",
          marker: "0.50",
          ttlmarker: "601.88",
          end: "1.38",
          plug1: "1.67",
          allplug1: "0.03",
          plug2: "1.64",
          allplug2: "0.03",
          plug12: "1.68",
        },
      ],
      rowc: [
        {
          defect: "",
          defectper: "",
          totalcutout: "",
          totalcutoutper: "",
          remnantsloss: "",
          remnantslossper: "",
          totallenthloss: "",
          totallenthlossper: "",
          recut: "",
        },
      ],
      rowz: [],
      cpartno: "",
      rowabc: [],
      rowxy: [],
      posts: [],
      posted:[],
      line_no: [],
      SO_NO_DOC: "",
      mu_glue_side: [],

      rowmain: [{}],

      rowsec: [
        {
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
        },
      ],

      rowDetail: [],
      isDisabled: false,
      isDisabled2: false,
      save_status: false,
      save_status_glue_side: false,
      category$clearable: false,
      category$clear: [],
      mu_seq: "",
      status_splice: false,
      save_status_splice: false,
      splice_no: [],
      data_top_qty_cut: [],

      exrow: [],
      re: [],
      date: "",
      GMT_TYPE: "",
      MARKER: "",
      F_TYPE: "",
    };
  },
 
  methods: {
    gmselect(index) {
      this.count_row_gm=1
      if (this.org == "" || this.F_TYPE == "") {
        this.$q.notify({
          message: "ORG or F TYPE No have Data",
          color: "green-9",
        });
      } else {
        if (this.rowx.gm !== "" && this.org !== "") {
          this.isDisabled2 = true;
        }
        const params = new FormData();
        params.append("GM_NO", this.rowx[index].gm_no);
        params.append("SO_YEAR", this.SO_YEAR);
        params.append("SO_NO", this.SO_NO);
        params.append("CPART_NO", this.cpartno);
        axios({
          method: "post",
          url: this.$api_url + "/somain.php/checkduplicate",
          data: params,
        }).then((resp) => {
          if (resp.data.data.length > 0) {
            this.$q.notify({
              position: "center",
              timeout: 1000,
              message: "Duplicate Data Please Try Again",
              color: "red-9",
            });

            this.rowx[index].gm_no = "";
          } else {
            axios({
              method: "post",
              url: this.$api_url + "/somain.php/check_detail",
              data: params,
            })
              .then((resp) => {
                console.log(resp.data);
                if (resp.data.data.length > 0) {
                  resp.data.data.forEach((e) => {
                    for (let xx = 0; xx <= this.rowx.length; xx++) {
                      if (xx == index) {
                        this.rowx[index].gm_no = e.GM_NO; 
                        this.rowx[index].GM_M2 = e.GM_M2;
                        this.rowx[index].GM_YD = e.GM_YD;
                        this.rowx[index].WIDTH_INCH = e.WIDTH_INCH;
                        this.rowx[index].LENGTH_YD = e.LENGTH_YD;
                        this.rowx[index].YD = e.YD;
                        this.rowx[index].LENGTH_INCH = e.LENGTH_INCH;
                        this.rowx[index].ttlmarker = parseFloat(
                          e.YD * e.LAYER
                        ).toFixed(2);
                        this.rowx[index].PER_U = e.PER_U;
                        this.rowx[index].LAYER = e.LAYER;
                        console.log(xx);
                        if (xx > 0) {
                          if (this.rowx[xx].GM_M2 !== this.rowx[0].GM_M2) {
                            this.$q.notify({
                              message: "ข้อมูล GM_M2 ไม่ตรงกัน",

                              color: "red-9",
                            });
                          }

                          if (this.rowx[xx].GM_YD !== this.rowx[0].GM_YD) {
                            this.$q.notify({
                              message: "ข้อมูล GM_YD ไม่ตรงกัน",

                              color: "red-9",
                            });
                          }

                          if (
                            this.rowx[xx].WIDTH_INCH !== this.rowx[0].WIDTH_INCH
                          ) {
                            this.$q.notify({
                              message: "ข้อมูล Weight Inch ไม่ตรงกัน",

                              color: "red-9",
                            });
                          }

                          if (this.rowx[xx].YD !== this.rowx[0].YD) {
                            this.$q.notify({
                              message: "ข้อมูล Length_YD(1) ไม่ตรงกัน",

                              color: "red-9",
                            });
                          }

                          if (
                            this.rowx[xx].LENGTH_YD !== this.rowx[0].LENGTH_YD
                          ) {
                            this.$q.notify({
                              message: "ข้อมูล LENGTH_YD(2) ไม่ตรงกัน",

                              color: "red-9",
                            });
                          }

                          if (
                            this.rowx[xx].LENGTH_INCH !==
                            this.rowx[0].LENGTH_INCH
                          ) {
                            this.$q.notify({
                              message: "ข้อมูล LENGTH_INCH ไม่ตรงกัน",

                              color: "red-9",
                            });
                          }

                          if (
                            this.rowx[xx].ttlmarker !== this.rowx[0].ttlmarker
                          ) {
                            this.$q.notify({
                              message: "ข้อมูล ttlmarker ไม่ตรงกัน",

                              color: "red-9",
                            });
                          }

                          if (this.rowx[xx].PER_U !== this.rowx[0].PER_U) {
                            this.$q.notify({
                              message: "ข้อมูล Marker ไม่ตรงกัน",

                              color: "red-9",
                            });
                          }

                          if (this.rowx[xx].LAYER !== this.rowx[0].LAYER) {
                            this.$q.notify({
                              message: "ข้อมูล PLY ไม่ตรงกัน",

                              color: "red-9",
                            });
                          }
                        }
                      }
                    }
                  });
                }
              })
              .catch((error) => {
                console.log(error);
              });
          }
        });
      
        console.log(index)

        const params2 = new FormData();
        params2.append("GM_NO", this.rowx[index].gm_no);
        params2.append("SO_YEAR", this.SO_YEAR);
        params2.append("SO_NO", this.SO_NO);

        axios({
          method: "post",
          url: this.$api_url + "/somain.php/check_qty_cut",
          data: params2,
        })
          .then((resp) => {
            console.log(resp.data);
            this.data_top_qty_cut = resp.data.data[0];
          })
          .catch((error) => {
            console.log(error);
          });

        this.exrow = this.rowx[index].gm_no;

        /*  if(value !== undefined)
        this.test.push(value)
        console.log(this.test)
        console.log(value) */
      }
    },
    clearall() {
      location.reload();
      /*   this.SO_YEAR = "";
      this.SO_NO = "";
      this.posts.SO_NO_DOC = "";
      this.posts.STYLE_REF = "";
      this.posts.CUST_NAME = "";
      this.posts.QTY = "";
      this.posts.GMT_TYPE = "";
      this.posts.FABRIC_TYPE = "";
      this.cpartno = [];
      this.rowx = [
        {
          name: "",
          number: "",
          weightm2: "",
          weightyd: "",
          widthinch: "",
          lenthyd: "",
          lenthyds: "",
          lenthinch: "",
          ttlmarker: "",
          marker: "",
          ply: "",
          issue: "",
          avgwidth: "",
          model: "",
        },
      ]; */
    },
    addRow() {

      console.log(this.rowwidth.length)
       this.rowwidth.push([
        {
          no: "",
          issue_roll:"",
          issue: "", 
          avg_width: "",
          delete: "",
        },
      ]);
    },
    addRowsplice() {
      if (this.row_splice.length > 3) {
        this.$q.notify({
          message: "Can't Add row more than 4 rows",

          color: "green-9",
        });
      } else {
        this.row_splice.push([
          {
            no: "",
            splice_value: "",
          },
        ]);
      }
    },
    addRowgm() {
      
      try{
        console.log(this.re[0].mu_seq)
      }catch(err){
          this.$q.notify({
          message: "Please Insert Data In First Record",
          color: "red-9",
        });
        this.status_add_row = false
      }
      console.log(this.status_add_row)
      if(this.status_add_row == true){
      this.count_row_gm++
      console.log(this.count_row_gm)
      this.rowx.push({
        name: "",
        weightm2: "",
        weightyd: "",
        widthinch: "",
        lenthyd: "",
        lenthyds: "",
        lenthinch: "",
        ttlmarker: "",
        marker: "",
        ply: "",
        number: "",
        issue: "",
        avgwidth: "",
        model: "",
        status: false,
      });
      this.save_status = false,
      console.log(this.rowx);
      }
      
    },
    removeRow() {
      this.rowwidth.pop();
    },
    Removegm() {
      this.rowx.pop();
    },
    deleteSelected(index) {
      console.log(index);
      console.log(this.rowwidth[index].no);
      console.log(this.exrow);
      console.log(this.re[0].mu_seq);
      const params = new FormData();
      params.append("MU_SEQ", this.re[0].mu_seq);
      params.append("GM_NO", this.exrow);
      params.append("LINE_NO", this.rowwidth[index].no);
      this.rowwidth.splice(index, 1);

      axios({
        method: "post",
        url: this.$api_url + "/detail.php/delete_sub_detail",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
        })
        .catch((error) => {
          console.log(error);
        });
    },
    deleteRow(index) {
      this.confirm_delete_detail = true
      this.delete_index = index
    },
    save() {
      console.log("success");
    },
    search() {
      this.row_fabric = []
      let username = this.$q.localStorage.getItem("username");
      
      if (this.SO_YEAR == null || this.SO_NO == "") {
        this.$q.notify({
          message: "Please Input SO_YEAR and SO_NO",
          color: "green-9",
        });
      } else {
        if (this.SO_YEAR.length > 2) {
          this.$q.notify({
            message: "Please user 2 Characters For SO YEAR",
            color: "red-9",
          });
        }
        if (this.SO_NO.length > 4) {
          this.$q.notify({
            message: "Please user 4 Characters For SO NO",
            color: "red-9",
          });
        } else {
          const params = new FormData();
          params.append("SO_YEAR", this.SO_YEAR);
          params.append("SO_NO", this.SO_NO);

          axios({
            method: "post",
            url: this.$api_url + "/somain.php/checkso",
            data: params,
          })
            .then((resp) => {
              console.log(resp.data)
              this.posts = resp.data.data[0];
             
            })
            .catch((error) => {
              console.log(error);
            });

     


          axios({
            method: "post",
            url: this.$api_url + "/somain.php/checkcpart",
            data: params,
          })
            .then((resp) => {
              var i = 0;

              resp.data.data.forEach((e) => {
                this.rowz[i] = e.CPART_NO;
                i++;
              });
            })
            .catch((error) => {
              console.log(error);
            });
          this.isDisabled = true;
        }
      }
    },
    checkgm() {
      const params = new FormData();
      params.append("SO_YEAR", this.SO_YEAR);
      params.append("SO_NO", this.SO_NO);
      params.append("cpartno", this.cpartno);
      this.rowxy = [];
      axios({
        method: "post",
        url: this.$api_url + "/somain.php/checkgm",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);

          var k = 0;

          resp.data.data.forEach((e) => {
            this.rowxy[k] = e.GM_NO; //นำค่าเข้าสู่ rowz
            k++;
          });
          console.log(this.rowxy);
        })
        .catch((error) => {
          console.log(error);
        })
        .finally(() => {});

         axios({
            method: "post",
            url: this.$api_url + "/somain.php/checkso2",
            data: params,
          })
            .then((resp) => {
              this.posted = resp.data.data[0];
              console.log(resp.data)
            })
            .catch((error) => {
              console.log(error);
            });

          const params2 = new FormData();
          params2.append("SO_YEAR", this.SO_YEAR);
          params2.append("SO_NO", this.SO_NO);
          params2.append("cpartno", this.cpartno);

           axios({
            method: "post",
            url: this.$api_url + "/somain.php/check_fabric",
            data: params2,
          })
          .then((resp)=>{
            console.log(resp.data)
             var x = 0;

              resp.data.data.forEach((e) => {
                this.row_fabric[x] = e.ITEM_DESC;
                x++;
              });
          })

    },
    addDetail(index) {
      this.count_layer = []
      console.log(this.start)
      this.isDisabled3 = true;
      this.check_status_wrong = true;
    
      this.exrow = this.rowx[index].gm_no
      this.save_status = false
      let username = this.$q.localStorage.getItem("username");

      if (this.exrow == "" || this.rowx[index].number == "" || this.fabric_type == "") {
        this.$q.notify({
          position: "center",
          timeout: 1000,
          size: "xl",
          message: "Please Insert Data Check FTYPE , Cpart , Fabric Type",
          color: "red-9",
        });
      } else {
        this.rowwidth = [
          {
            no: "1",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "2",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "3",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "4",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "5",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "6",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "7",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "8",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "9",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "10",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "11",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "12",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "13",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "14",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "15",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "16",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "17",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "18",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "19",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "20",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
        ];
       
        this.gm_no = this.rowx[index].gm_no;
        this.table = this.rowx[index].number;
        this.weightm2 = this.rowx[index].GM_M2;
        this.weightyd = this.rowx[index].GM_YD;
        this.widthinch = this.rowx[index].WIDTH_INCH;
        this.lenthyds = this.rowx[index].YD;
        this.lenthyd = this.rowx[index].LENGTH_YD;
        this.lenthinch = this.rowx[index].LENGTH_INCH;
        this.rowx[index].ttlmarker = parseFloat(
          this.rowx[index].YD * this.rowx[index].LAYER
        ).toFixed(2);
        this.ttlmarker = this.rowx[index].ttlmarker;
        this.marker = this.rowx[index].PER_U;
        this.ply = this.rowx[index].LAYER;
        console.log(index)
        for(var ax = 0; ax<index+1; ax++){
            this.count_layer.push({
          value:this.rowx[index].LAYER
        })
        }
        
        if(this.weightm2 == 0){
            this.$q.notify({
            position: "center",
            timeout: 800,
            size: "xl",
            message: "มีค่าข้อมูลผิดพลาด โปรดแจ้งกับทาง Programmer ให้ตรวจสอบ",
            color: "red-9",
          });
          this.check_status_wrong = false
        }

        if(this.weightyd == 0){
            this.$q.notify({
            position: "center",
            timeout: 800,
            size: "xl",
            message: "มีค่าข้อมูลผิดพลาด โปรดแจ้งกับทาง Programmer ให้ตรวจสอบ",
            color: "red-9",
          });
          this.check_status_wrong = false
        }

        if(this.widthinch == 0){
            this.$q.notify({
            position: "center",
            timeout: 800,
            size: "xl",
            message: "มีค่าข้อมูลผิดพลาด โปรดแจ้งกับทาง Programmer ให้ตรวจสอบ",
            color: "red-9",
          });
          this.check_status_wrong = false
        }

        if(this.lenthyds == 0){
            this.$q.notify({
            position: "center",
            timeout: 800,
            size: "xl",
            message: "มีค่าข้อมูลผิดพลาด โปรดแจ้งกับทาง Programmer ให้ตรวจสอบ",
            color: "red-9",
          });
          this.check_status_wrong = false
        }

        if(this.lenthinch == 0){
            this.$q.notify({
            position: "center",
            timeout: 800,
            size: "xl",
            message: "มีค่าข้อมูลผิดพลาด โปรดแจ้งกับทาง Programmer ให้ตรวจสอบ",
            color: "red-9",
          });
          this.check_status_wrong = false
        }

        if(this.ply == 0){
            this.$q.notify({
            position: "center",
            timeout: 800,
            size: "xl",
            message: "มีค่าข้อมูลผิดพลาด โปรดแจ้งกับทาง Programmer ให้ตรวจสอบ",
            color: "red-9",
          });
          this.check_status_wrong = false
        }

     if(this.check_status_wrong == true){  

        if (this.status_check_gm == true) {
          this.$q.notify({
            position: "center",
            timeout: 800,
            size: "xl",
            message: "โปรดตั้งค่าข้อมูลในแต่ละ GM ให้ตรงกัน",
            color: "red-9",
          });
        } else {
          this.basic = true;
          console.log(this.gm_no);
          if (this.re == "") {
            axios.get(this.$api_url + "/somain.php/createmu")
            .then((resp) => {
              this.re.push({
                mu_seq: resp.data.data[0].NEXTVAL,
              });

              if (this.F_TYPE == "1. WOVEN - NYLON, TAFFETA") {
                this.F_TYPE = 1;
              }
              if (
                this.F_TYPE ==
                "2. LIGHT-MIDDLE WEIGHT - SINGLE JERSEY, INTERLOCK, DOUBLE KNIT, JACQUARD"
              ) {
                this.F_TYPE = 2;
              }
              if (this.F_TYPE == "3. FLEECE") {
                this.F_TYPE = 3;
              }
              if (
                this.F_TYPE ==
                "4. SPECIAL - RIB, ELASTIC, ELASTAN, LYCRA, SPANDEX, TWO WAY, WAFFLE, PRE WASH, PRESHRUNK, VELOUR"
              ) {
                this.F_TYPE = 4;
              }
              
              
        
  
        
              if (this.re !== "") {
                const params = new FormData();
                params.append("mu_seq", this.re[0].mu_seq);
                params.append("GM_NO", this.gm_no);
                params.append("table", this.table);
                params.append("weightm2", this.weightm2);
                params.append("weightyd", this.weightyd);
                params.append("widthinch", this.widthinch);
                params.append("lenthyd", this.lenthyd);
                params.append("lenthyds", this.lenthyds);
                params.append("lenthinch", this.lenthinch);
                params.append("ttlmarker", this.ttlmarker);
                params.append("marker", this.marker);
                params.append("ply", this.ply);
                params.append("SO_YEAR", this.SO_YEAR);
                params.append("SO_NO", this.SO_NO);
                params.append("SO_NO_DOC", this.posts.SO_NO_DOC);
                params.append("STYLE", this.posts.STYLE_REF);
                params.append("CUSTOMER", this.posts.CUST_NAME);
                params.append("QTY", this.posts.QTY);
                params.append("GM_TYPE", this.posts.GMT_TYPE);
                params.append("F_TYPE", this.F_TYPE);
                if(this.fabric_type == null){
                  params.append("FABRIC_TYPE",this.posted.ITEM_DESC)
                }else{
                params.append("FABRIC_TYPE", this.fabric_type);
                }
                params.append("DATE", this.date);
                params.append("CPART_NO", this.cpartno),
                params.append("ORGE", this.org),
                params.append("ENDLOSS_AUTO", this.endloss_auto),
                params.append("ENDLOSS_AUTO_INCH", this.endloss_auto_inch),
                params.append("USERNAME", username);
                let day = this.start.substring(0, 2);
                let month = this.start.substring(3, 5);
                let year = this.start.substring(6, 10);
                this.start_date = year +"/"+ month +"/"+ day
                params.append("DATE_MU_CREATE",this.start_date)
                for (var pair of params.entries()) {
                  console.log(pair[0] + ", " + pair[1]);
                }

               
                axios({
                  method: "post",
                  url: this.$api_url + "/somain.php/create_for_head",
                  data: params,
                })
                  .then((resp) => {
                    console.log(resp.data);
                    if (resp.data.status == true) {
                      axios({
                        method: "post",
                        url: this.$api_url + "/somain.php/createdetail",
                        data: params,
                      })
                        .then((resp) => {
                          console.log(resp.data)
                          this.stat = resp.data.status;

                          if (this.stat == true) {
                            this.rowx[index].status = true;
                            this.save_status = false;
                            this.status_check_gm = false;
                          }
                        })
                        .catch((error) => {
                          console.log(error);
                        });
                    } else {
                    }
                  })
                  .catch((error) => {
                    console.log(error);
                  });
              }
              
            });
            this.result1 =
              Number(this.endloss_auto * 36) + Number(this.endloss_auto_inch);
            console.log(this.result1);
            console.log(this.rowx[0].LENGTH_YD);
            console.log(this.rowx[0].LENGTH_INCH);
            this.result2 =
              Number(this.rowx[0].LENGTH_YD * 36) +
              Number(this.rowx[0].LENGTH_INCH);
            console.log(this.result2);
            this.result =
              Number(this.result1) -
              (Number(this.result2) * this.rowmain[0].ply) / 36;
            console.log(this.result);
            this.rowmain[0].avgglueside = this.result;
            console.log(this.rowmain[0].avgglueside);
          } else {
            const params = new FormData();
            params.append("mu_seq", this.re[0].mu_seq);
            params.append("GM_NO", this.gm_no);
            params.append("table", this.table);
            params.append("weightm2", this.weightm2);
            params.append("weightyd", this.weightyd);
            params.append("widthinch", this.widthinch);
            params.append("lenthyd", this.lenthyd);
            params.append("lenthyds", this.lenthyds);
            params.append("lenthinch", this.lenthinch);
            params.append("ttlmarker", this.ttlmarker);
            params.append("marker", this.marker);
            params.append("ply", this.ply);
            params.append("SO_YEAR", this.SO_YEAR);
            params.append("SO_NO", this.SO_NO);
            params.append("SO_NO_DOC", this.posts.SO_NO_DOC);
            params.append("STYLE", this.posts.STYLE_REF);
            params.append("CUSTOMER", this.posts.CUST_NAME);
            params.append("QTY", this.posts.QTY);
            params.append("GM_TYPE", this.posts.GMT_TYPE);
            
            if(this.fabric_type == null){
                  params.append("FABRIC_TYPE",this.posted.ITEM_DESC)
                }else{
                params.append("FABRIC_TYPE", this.fabric_type);
                }
            params.append("DATE", this.date);
            console.log(this.start)
            let day = this.start.substring(0, 2);
            let month = this.start.substring(3, 5);
            let year = this.start.substring(6, 10);
            this.start_date = year +"/"+ month +"/"+ day
           
            params.append("DATE_MU_CREATE",this.start_date )
            params.append("CPART_NO", this.cpartno);
            params.append("username", username);
            for (var pair of params.entries()) {
              console.log(pair[0] + ", " + pair[1]);
            }
            axios({
              method: "post",
              url: this.$api_url + "/somain.php/createdetail",
              data: params,
            })
              .then((resp) => {
                console.log(resp.data)
                this.stat = resp.data.status;

                if (this.stat == true) {
                  this.rowx[index].status = true;
                  this.status_check_gm = false;
                }
              })
              .catch((error) => {
                console.log(error);
              });
          }
        }
     }
        console.log(this.count_layer)
        this.sum_count_layer = 0
        for(var ax = 0; ax<this.count_layer.length; ax++){
        this.sum_count_layer = Number(this.sum_count_layer) + Number(this.count_layer[ax].value)
        }

        const params5 = new FormData();
        params5.append("SO_NO",this.SO_NO)
        params5.append("SO_YEAR",this.SO_YEAR)
        params5.append("GM_NO",this.gm_no)
        params5.append("PLY", this.sum_count_layer);

        axios({
          method: "post",
          url: this.$api_url + "/somain.php/check_qty_cut_2",
          data: params5,
        })
          .then((resp) => {
            console.log(resp.data);
            this.data_top_qty_cut = resp.data.data[0];
          })
          .catch((error) => {
            console.log(error);
          });

        this.result1 =
          Number(this.endloss_auto * 36) + Number(this.endloss_auto_inch);
        console.log(this.result1);
        console.log(this.rowx[0].LENGTH_YD);
        console.log(this.rowx[0].LENGTH_INCH);
        this.result2 =
          Number(this.rowx[0].LENGTH_YD * 36) +
          Number(this.rowx[0].LENGTH_INCH);
        console.log(this.result2);
        this.result =
          Number(this.result1) -
          (Number(this.result2) * this.rowmain[0].ply) / 36;
        console.log(this.result);
        this.rowmain[0].avgglueside = this.result;
        console.log(this.rowmain[0].avgglueside);
      }
      
    },
    async goDetail(index) {
      this.rowexport=[]
      this.exrow = this.rowx[index].gm_no
      this.save_status = true
      const params2 = new FormData();
      params2.append("mu_seq", this.re[0].mu_seq);
      params2.append("GM_NO", this.rowx[index].gm_no);
      params2.append("table", this.rowx[index].number);
      params2.append("weightm2", this.rowx[index].GM_M2);
      params2.append("weightyd", this.rowx[index].GM_YD);
      params2.append("widthinch", this.rowx[index].WIDTH_INCH);
      params2.append("lenthyd", this.rowx[index].LENGTH_YD);
      this.rowx[index].YD = parseFloat(
        Number(this.rowx[index].LENGTH_INCH / 36) +
          Number(this.rowx[index].LENGTH_YD)
      ).toFixed(2);
      params2.append("lenthyds", this.rowx[index].YD);
      params2.append("lenthinch", this.rowx[index].LENGTH_INCH);
      this.rowx[index].ttlmarker = parseFloat(
        this.rowx[index].YD * this.rowx[index].LAYER
      ).toFixed(2);
      params2.append("ttlmarker", this.rowx[index].ttlmarker);
      params2.append("marker", this.rowx[index].PER_U);
      params2.append("ply", this.rowx[index].LAYER);
      let username = this.$q.localStorage.getItem("username");
      params2.append("username", username);
      for (var pair of params2.entries()) {
        console.log(pair[0] + ", " + pair[1]);
      }
      

    await  axios({
        method: "post",
        url: this.$api_url + "/somain.php/update_detail",
        data: params2,
      })
        .then((resp) => {
          console.log(resp.data);
        })
        .catch((error) => {
          console.log(error);
        });

      this.rowwidth = [];
      this.basic = true;

      this.exrow = this.rowx[index].gm_no;
      console.log(this.exrow);

      const params = new FormData();

      params.append("GM_NO", this.exrow);
      params.append("MU_SEQ", this.re[0].mu_seq);
     await  axios({
        method: "post",
        url: this.$api_url + "/somain.php/returndata",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data.data);
          if (resp.data.data.length > 0) {
            resp.data.data.forEach((e) => {
              this.rowwidth.push({
                no: e.LINE_NO,
                issue_kg: e.ISSUE,
                issue_yard: e.ISSUE_YRD,
                avg_width: e.AVG_WIDTH,

                delete: "",
              }),
                console.log(this.rowwidth);
            });
          } else {
            this.save_status = false;
            this.rowwidth = [
          {
            no: "1",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "2",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "3",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "4",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "5",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "6",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "7",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "8",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "9",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "10",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "11",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "12",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "13",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "14",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "15",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "16",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "17",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "18",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "19",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "20",
            issue_yard: "",
            issue_kg: "",
            avg_width: "",
            delete: "",
          },
        ];
          }
          this.result1 =
            Number(this.endloss_auto * 36) + Number(this.endloss_auto_inch);
      
          this.result2 =
            Number(this.rowx[0].LENGTH_YD * 36) +
            Number(this.rowx[0].LENGTH_INCH);
          
          this.result =
            Number(this.result1) -
            (Number(this.result2) * this.rowmain[0].ply) / 36;
         
          this.rowmain[0].avgglueside = this.result;
          this.check_bottom_data();
        })
        .catch((error) => {
          console.log(error);
        });
      
        console.log(this.count_layer)
        for(var ax = 0; ax<this.count_layer.length; ax++){
          console.log(this.count_layer[ax].value)
        }
        
        if(this.re[0].mu_seq !== ""){
        const params3 = new FormData();
        params3.append("MU_SEQ",this.re[0].mu_seq)

        await axios({
          method: "post",
          url: this.$api_url + "/somain.php/keep_data_for_ply",
          data: params3,
        }).then((resp)=>{
           if (resp.data.data.length > 0) {
            resp.data.data.forEach((e) => {
              this.rowexport.push({
                mu_seq:e.MU_SEQ,
                gm_no:e.GM_NO,
                ply:e.PLY
              });
            });
          }
        })
        }
        this.sum_value_ply = 0  
        for(var ax =0; ax<this.rowexport.length; ax++){
        this.sum_value_ply = Number(this.sum_value_ply) + Number(this.rowexport[ax].ply)
        }
         const params5 = new FormData();
        params5.append("SO_NO",this.SO_NO)
        params5.append("SO_YEAR",this.SO_YEAR)
        params5.append("GM_NO",this.gm_no)
        params5.append("PLY", this.sum_value_ply);

     
        

     await   axios({
          method: "post",
          url: this.$api_url + "/somain.php/check_qty_cut_2",
          data: params5,
        })
          .then((resp) => {
            console.log(resp.data);
            this.data_top_qty_cut = resp.data.data[0];
          })
          .catch((error) => {
            console.log(error);
          });



    },

    Savedata() {
      
      
       let countrow = this.rowwidth.length;
      for (var x = 1; x <= this.rowwidth.length; x++) {
        if (this.rowwidth[x - 1].issue_kg == "" && this.rowwidth[x - 1].avg_width == "") {
          countrow = countrow - 1;
        }
      }
     
      var count = 0;
      for (var i = 0; i < countrow; i++) {
        const params = new FormData();
        
          params.append("MU_SEQ", this.re[0].mu_seq);
          params.append("GM_NO", this.exrow);
          if (this.rowwidth[i].issue_kg == undefined) {
            this.rowwidth[i].issue_kg = "";
          }
          params.append("ISSUE", this.rowwidth[i].issue_kg);
          if (this.rowwidth[i].issue_yard == undefined) {
            this.rowwidth[i].issue_yard = "";
          }
          params.append("ISSUE_YARD", this.rowwidth[i].issue_yard);
          if (this.rowwidth[i].avg_width == undefined) {
            this.rowwidth[i].avg_width = "";
          }
          params.append("AVG_WIDTH", this.rowwidth[i].avg_width);
          axios
            .get(this.$api_url + "/somain.php/create_line_no")
            .then((resp) => {
              this.line_no = resp.data.data[0].NEXTVAL;

              params.append("LINE_NO", this.line_no);
           
              axios({
                method: "post",
                url: this.$api_url + "/somain.php/createsubdetail",
                data: params,
              })
                .then((resp) => {
               
                  if (resp.data.status == true) {
                    this.save_status = true;
                    this.$q.notify({
                      message:
                        "Successful" +
                        " " +
                        "GM NO:" +
                        this.exrow +
                        " " +
                        "MU SEQ:" +
                        this.re[0].mu_seq,
                      color: "green-9",
                      position: "center",
                    });
                    if (count == countrow - 1) {
                      this.check_bottom_data();
                     
                    }
                    count++;
                  }
                })
                .catch((error) => {
                  console.log(error);
                }); 
            });
        
      } 
    },
    change_cause() {
      

      if (this.rowmain[0].code == "A มาร์คไม่มีผล") {
        this.rowmain[0].cause = "มาร์คไม่มีผล";
        this.rowmain[0].code = "A";
       
      }

      if (this.rowmain[0].code == "-") {
        this.rowmain[0].cause = "";
        this.rowmain[0].code = "";
 
      }
      if (this.rowmain[0].code == "B ไม่ได้คัดแยกหน้าผ้าปู") {
        this.rowmain[0].cause = "ไม่ได้คัดแยกหน้าผ้าปู";
        this.rowmain[0].code = "B";
       
      }
      if (this.rowmain[0].code == "C หน้าผ้าไม่สม่ำเสมอในม้วน และต่างม้วน") {
        this.rowmain[0].cause = "หน้าผ้าไม่สม่ำเสมอในม้วน และต่างม้วน";
        this.rowmain[0].code = "C";
       
      }
      if (this.rowmain[0].code == "D ใช้มาร์คไม่เหมาะสมกับหน้าผ้า") {
        this.rowmain[0].cause = "ใช้มาร์คไม่เหมาะสมกับหน้าผ้า";
        this.rowmain[0].code = "D";
        
      }
      if (
        this.rowmain[0].code ==
        "E อื่น ๆ ริมเสียบางช่วง, ริมย้วย, ปูผ้าหลายสีหน้าผ้าไม่เท่ากัน"
      ) {
        this.rowmain[0].cause =
          "อื่น ๆ ริมเสียบางช่วง, ริมย้วย, ปูผ้าหลายสีหน้าผ้าไม่เท่ากัน";
        this.rowmain[0].code = "E";
        
      }
    },
    saveAll() {
      if (this.org == "") {
        this.$q.notify({
          message: "Please Input ORG",
          color: "green-9",
        });
      } else {
        const params = new FormData();

        console.log(this.re[0].mu_seq);
        this.mu_doc_no = this.re[0].mu_seq + this.SO_YEAR + this.SO_NO;
        console.log(this.rowsec[0].code);
        params.append("MU_SEQ", this.re[0].mu_seq);

        params.append("SO_YEAR", this.SO_YEAR);
        params.append("SO_NO", this.SO_NO);
        params.append("SO_NO_DOC", this.posts.SO_NO_DOC);

        if (this.rowmain[0].endplug1 == undefined) {
          this.rowmain[0].endplug1 = "";
        }
        if (this.rowmain[0].endplug2 == undefined) {
          this.rowmain[0].endplug2 = "";
        }
        if (this.rowmain[0].code == undefined) {
          this.rowmain[0].code = "";
        }
        if (this.rowmain[0].cause == undefined) {
          this.rowmain[0].cause = "";
        }
        if (this.rowmain[0].splicenoof == null) {
          this.rowmain[0].splicenoof = "";
        }
        if (
          this.rowmain[0].spliceinch == undefined ||
          this.rowmain[0].spliceinch == "NaN"
        ) {
          this.rowmain[0].spliceinch = "";
        }

        if (this.rowmain[0].splicenoof == null) {
          this.rowmain[0].splicenoof = "";
        }
        if (this.rowmain[0].cutoutnoof == undefined) {
          this.rowmain[0].cutoutnoof = "";
        }
        if (this.rowmain[0].cutoutweight == undefined) {
          this.rowmain[0].cutoutweight = "";
        }
        if (this.rowmain[0].remnants == undefined) {
          this.rowmain[0].remnants = "";
        }
        if (this.rowmain[0].widthplug1 == undefined) {
          this.rowmain[0].widthplug1 = "";
        }
        if (this.rowmain[0].widthplug2 == undefined) {
          this.rowmain[0].widthplug2 = "";
        }
        if (
          this.rowmain[0].avgglueside == undefined ||
          this.rowmain[0].avgglueside == "NaN"
        ) {
          this.rowmain[0].avgglueside = "";
        }
        if (
          this.rowmain[0].re_cut == undefined ||
          this.rowmain[0].re_cut == "NaN"
        ) {
          this.rowmain[0].re_cut = "";
        }
        if (
          this.rowmain[0].endplug12_yd == undefined ||
          this.rowmain[0].endplug12_yd == "NaN"
        ) {
          this.rowmain[0].endplug12_yd = "";
        }
        if (
          this.rowmain[0].plug12 == undefined ||
          this.rowmain[0].plug12  == "NaN"
        ) {
          this.rowmain[0].plug12  = "";
        }

        if (this.rowmain[0].code == "A มาร์คไม่มีผล") {
          this.rowmain[0].code = "A";
          
        }

        if (this.rowmain[0].code == "-") {
          this.rowmain[0].cause = "";
         
        }
        if (this.rowmain[0].code == "B ไม่ได้คัดแยกหน้าผ้าปู") {
          this.rowmain[0].cause = "B";
          
        }
        if (this.rowmain[0].code == "C หน้าผ้าไม่สม่ำเสมอในม้วน และต่างม้วน") {
          this.rowmain[0].cause = "C";
         
        }
        if (this.rowmain[0].code == "D ใช้มาร์คไม่เหมาะสมกับหน้าผ้า") {
          this.rowmain[0].cause = "D";
          
        }
        if (
          this.rowmain[0].code ==
          "E อื่น ๆ ริมเสียบางช่วง, ริมย้วย, ปูผ้าหลายสีหน้าผ้าไม่เท่ากัน"
        ) {
          this.rowmain[0].cause = "E";
         
        }

        
        if(this.F_TYPE == "1. WOVEN - NYLON, TAFFETA")
        {
        this.F_TYPE = 1  
        }

        if(this.F_TYPE == "2. LIGHT-MIDDLE WEIGHT - SINGLE JERSEY, INTERLOCK, DOUBLE KNIT, JACQUARD")
        {
        this.F_TYPE = 2  
        }

        if(this.F_TYPE == "3. FLEECE")
        {
        this.F_TYPE = 3 
        }

        if(this.F_TYPE == "4. SPECIAL - RIB, ELASTIC, ELASTAN, LYCRA, SPANDEX, TWO WAY, WAFFLE, PRE WASH, PRESHRUNK, VELOUR")
        {
        this.F_TYPE = 4  
        }
        

        params.append("start",this.start)
        params.append("f_type",this.F_TYPE)
        params.append("END_PLUG1", this.rowmain[0].endplug1);
        params.append("END_PLUG2", this.rowmain[0].endplug2);
        params.append("CAUSE_CODE", this.rowmain[0].code);
        params.append("CAUSE_NAME", this.rowmain[0].cause);
        params.append("SPLICE_NOOFF", this.rowmain[0].splicenoof);
        params.append("SPICE_INCH", this.rowmain[0].spliceinch);
        params.append("CUT_OUT", 0);
        params.append("CUT_OUT_GRAM", this.rowmain[0].cutoutweight);
        params.append("REMNANTS_WEIGHT", this.rowmain[0].remnants);
        params.append("PLUG1_GRAM", this.rowmain[0].widthplug1);
        params.append("PLUG2_GRAM", this.rowmain[0].widthplug2);
        params.append("GLUE_SIDE", this.rowmain[0].avgglueside);
        params.append("ENDLOSS_AUTO", this.endloss_auto);
        params.append("ENDLOSS_AUTO_INCH", this.endloss_auto_inch);
        params.append("END_PLUG1_2_YD", this.rowmain[0].endplug12_yd);
        params.append("PLUG12", this.rowmain[0].plug12);
        let username = this.$q.localStorage.getItem("username");
        params.append("username", username);
        params.append("RE_CUT", this.rowmain[0].re_cut);
        params.append("MU_DOC_NO", this.mu_doc_no);
        for (var pair of params.entries()) {
                  console.log(pair[0] + ", " + pair[1]);
                }
        

         axios({
          method: "post",
          url: this.$api_url + "/somain.php/update_head",
          data: params,
        })
          .then((resp) => {
            console.log(resp.data);

            this.rowmain = [];

            if (resp.data.status == true) {
              this.check_bottom_data();
            }
            //this.rowx[index].gm_no = "";
          })
          .catch((error) => {
            console.log(error);
          })
          .finally(() => {
            this.$q.notify({
              message: "Save Complete",
              color: "Black",
            });
            //this.rowx[index].gm_no = "";
          });
      }
    },
    update_detail_sub(index) {
      for (var i = 0; i < this.rowwidth.length; i++) {
        if (this.rowwidth[i].no !== undefined) {
         
          const params = new FormData();
          params.append("line_no", this.rowwidth[i].no);
          params.append("mu_seq", this.re[0].mu_seq);
          if (this.rowwidth[i].issue_yard == undefined || this.rowwidth[i].issue_yard == null) {
            this.rowwidth[i].issue_yard = "";
          }
          params.append("issue_yard", this.rowwidth[i].issue_yard);
          if (this.rowwidth[i].issue_kg == undefined || this.rowwidth[i].issue_kg == null) {
            this.rowwidth[i].issue_kg = "";
          }
          params.append("issue_kg", this.rowwidth[i].issue_kg);
          if (this.rowwidth[i].avg_width == undefined || this.rowwidth[i].avg_width == null) {
            this.rowwidth[i].avg_width = "";
          }
          params.append("avg_width", this.rowwidth[i].avg_width);
          let username = this.$q.localStorage.getItem("username");
          params.append("username", username);
          

          axios({
            method: "post",
            url: this.$api_url + "/somain.php/update_detail_sub",
            data: params,
          }).then((resp) => {
            console.log(resp.data);
            if (resp.data.status == true) {
              this.save_status = true;

              this.$q.notify({
                message:
                  "Update Sucessful" +
                  " " +
                  "GM NO:" +
                  this.exrow +
                  " " +
                  "MU SEQ:" +
                  this.re[0].mu_seq,
                color: "green-9",
              });
            }
          });
        } else {
          const params2 = new FormData();
          params2.append("MU_SEQ", this.re[0].mu_seq);
          params2.append("GM_NO", this.exrow);

          if (this.rowwidth[i].issue_yard == undefined || this.rowwidth[i].issue_yard == null) {
            this.rowwidth[i].issue_yard = "";
          }
          params2.append("ISSUE_YARD", this.rowwidth[i].issue_yard);
          if (this.rowwidth[i].issue_kg == undefined || this.rowwidth[i].issue_kg == null) {
            this.rowwidth[i].issue_kg = "";
          }
          params2.append("ISSUE", this.rowwidth[i].issue_kg);
          if (this.rowwidth[i].avg_width == undefined || this.rowwidth[i].avg_width == null) {
            this.rowwidth[i].avg_width = "";
          }
          params2.append("AVG_WIDTH", this.rowwidth[i].avg_width);
          let username = this.$q.localStorage.getItem("username");
          params2.append("username", username);

          // debugger
          axios
            .get(this.$api_url + "/somain.php/create_line_no")
            .then((resp) => {
              
              this.line_no = resp.data.data[0].NEXTVAL;

              params2.append("LINE_NO", this.line_no);
             

              axios({
                method: "post",
                url: this.$api_url + "/somain.php/createsubdetail",
                data: params2,
              })
                .then((resp) => {
               
                  if (resp.data.status == true) {
                    this.save_status = true;

                    this.rowwidth = [
                      {
                        no: "1",
                        issue_kg: "",
                        issue_yard: "",
                        avg_width: "",
                        delete: "",
                      },
                      {
                        no: "2",
                        issue_kg: "",
                        issue_yard: "",
                        avg_width: "",
                        delete: "",
                      },
                      {
                        no: "3",
                        issue_kg: "",
                        issue_yard: "",
                        avg_width: "",
                        delete: "",
                      },
                      {
                        no: "4",
                        issue_kg: "",
                        issue_yard: "",
                        avg_width: "",
                        delete: "",
                      },
                      {
                        no: "5",
                        issue_kg: "",
                        issue_yard: "",
                        avg_width: "",
                        delete: "",
                      },
                    ];
                  }
                })
                .catch((error) => {
                  console.log(error);
                });
            });
        }
      }
      this.rowmain = [];
      this.check_bottom_data();
    },
    check_bottom_data() {
      const params = new FormData();
      this.rowmain = [];
      params.append("MU_SEQ", this.re[0].mu_seq);
      params.append("SO_NO_DOC", this.posts.SO_NO_DOC);
      
      axios({
        //all detail ด้านล่าง
        method: "post",
        url: this.$api_url + "/detail.php/all_data_bottom",
        data: params,
      })
        .then((resp) => {
          if (resp.data.data.length > 0) {
            resp.data.data.forEach((e) => {
              this.rowmain.push({
               issue: e.ISSUE,
                issue_roll: e.ROLL_NO,
                avg_width: e.AVG_WIDTH,
                gmm2: e.WEIGHT_M2,
                gmyd: e.WEIGHT_YD,
                width_inch: e.WIDTH_INCH,
                length_ydb: e.YARD,
                length_inch: e.LENGTH_INCH,
                obs_width: e.OBS_WIDTH,
                marker: e.MARKER,
                ply: e.PLY,
                endplug1: e.END_PLUG1,
                endplug2: e.END_PLUG2,
                cause: e.CAUSE_NAME,
                code: e.CAUSE_CODE,
                splicenoof: e.SPLICE,
                spliceinch: parseFloat(e.SPLICE_INCH).toFixed(2),
                cutoutnoof: e.CUT_OUT,
                cutoutweight: e.CUT_OUT_GRAM,
                remnants: e.REMNANTS_WEIGHT,
                widthplug1: e.PLUG1_GRAM,
                widthplug2: e.PLUG2_GRAM,
                avgglueside: parseFloat(e.GLUE_SIDE).toFixed(3),
                ttl_marker: parseFloat(e.TTL_MARKER_YD).toFixed(2), //เอาสูตรออก
                yard_roll: e.YARD_ROLL,
                length_yd: e.LENGTH_YD,
                obswidth: e.OBS_WIDTH,
                ttlobs_widthloss: e.OBS_WIDTH,
                end1_2: e.END1_2,
                plug1_inch: e.PLUG1_INCH,
                plug2_inch: e.PLUG2_INCH,
                per_plug1: e.PER_PLUG1,
                per_plug2: e.PER_PLUG2,
                
                endloss_yd: parseFloat(
                  (6 *
                    (Number(e.END_PLUG1) + Number(e.END_PLUG2)) *
                    e.WIDTH_INCH) /
                    e.WEIGHT_YD /
                    36
                ).toFixed(2),
                widthloss: parseFloat(e.WITH_LOSS).toFixed(4),
                endless_length_yd: e.AVG_END_LOSS_YD,
                avg_end: parseFloat(e.AVG_END_LOSS).toFixed(4),
                sectionlosslenth: parseFloat(e.SECTION_LOSS_YD).toFixed(2),
                sectionlossper: e.SECTION_LOSS_PER,
                spliceloss: e.SPLICE_LOSS_YD,
                splicelossper: parseFloat(e.SPLICE_LOSS_PER).toFixed(4),
                overlap: e.OVER_LAB_YD,
                overlapper: parseFloat(e.OVER_LAB_PER).toFixed(2),
                defect: parseFloat(e.DEFECT_YD).toFixed(2),
                defectper: parseFloat(e.DEFECT_PER).toFixed(2),
                totalcutout: parseFloat(e.TOTAL_CUTOUT_YD).toFixed(2),
                totalcutoutper: parseFloat(e.TOTAL_CUTOUT_PER).toFixed(4),
                rement: parseFloat(e.REMENT_LOSS_YD).toFixed(2),
                rement_per: parseFloat(e.REMENT_LOSS_PER).toFixed(4),
                totallenthloss: parseFloat(e.TOTAL_LENGTH_LOSS_YD).toFixed(2),
                totallenthlossper: parseFloat(e.TOTAL_LEN_LOSS_PER).toFixed(2),
                endplug12_yd: parseFloat(e.END_PLUG1_2_YD).toFixed(2),
                re_cut: parseFloat(e.RE_CUT).toFixed(2),
                plug12_inch: parseFloat(e.PLUG1_2_INCH).toFixed(2),
                avg_end_code: e.AVG_END_CODE + " " + e.AVG_END_CAUSE,
                per_splice_loss_code:
                  e.SPLICE_LOSS_PER_CODE + " " + e.SPLICE_LOSS_PER_CAUSE,
                remnants_loss_code:
                  e.REMNANTS_LOSS_CODE + " " + e.REMNANTS_LOSS_CAUSE,
                total_cut_out_per_code:
                  e.TOTAL_CUT_OUT_PER_CODE + " " + e.TOTAL_CUT_OUT_PER_CAUSE,
                widthloss_code: e.WIDTH_LOSS_CODE + " " + e.WIDTH_LOSS_CAUSE,
                plug12:parseFloat(e.PLUG1_2_INCH_AM).toFixed(2),
              });
            });
          }
        })
         .catch((error) => {
          console.log(error);
        }); 
    },
    add_glue_side() {
      this.status_glue = true;
    },
    Save_data_glue_side() {
      this.result_glue_side1 = 0;
      this.result_glue_side2 = 0;
      this.mu_seq = 1515;

      
      //this.re[0].mu_seq

      for (var i = 0; i < 3; i++) {
        const params = new FormData();

        params.append("MU_SEQ", this.re[0].mu_seq);
        params.append("GLUE_SIDE1", this.row_glue_side[i].glue_side1);
        params.append("GLUE_SIDE2", this.row_glue_side[i].glue_side2);
        axios
          .get(this.$api_url + "/somain.php/createmu_glue_side")
          .then((resp) => {
            this.line_no = resp.data.data[0].NEXTVAL;

            params.append("glue_no", this.line_no);
            

            void axios({
              method: "post",
              url: this.$api_url + "/somain.php/save_data_glue_side",
              data: params,
            }).then((resp) => {
              console.log(resp.data);
            });
          });
        this.result_glue_side1 =
          Number(this.result_glue_side1) +
          Number(this.row_glue_side[i].glue_side1);
        this.result_glue_side2 =
          Number(this.result_glue_side2) +
          Number(this.row_glue_side[i].glue_side2);
      }

      this.result_glue_side1 = Number(this.result_glue_side1) / 3;
      this.result_glue_side2 = Number(this.result_glue_side2) / 3;
      if (this.result_glue_side1 !== "" && this.result_glue_side2 !== "") {
        this.rowmain[0].avgglueside =
          Number(this.result_glue_side1) + Number(this.result_glue_side2);
      }
      this.save_status_glue_side = true;
      this.status_glue = false;
    },
    check_glue_side() {
      this.row_glue_side = [];
      this.status_glue = true;
      this.mu_seq = 1515;

      const params2 = new FormData();
      params2.append("mu_seq", this.re[0].mu_seq);
      axios({
        method: "post",
        url: this.$api_url + "/somain.php/check_data_glue_side",
        data: params2,
      })
        .then((resp) => {
          
          if (resp.data.data.length > 0) {
            resp.data.data.forEach((e) => {
              this.row_glue_side.push({
                glue_no: e.GLUE_NO,
                glue_side1: e.GLUE_SIDE1,
                glue_side2: e.GLUE_SIDE2,
              }),
                console.log(this.row_glue_side)
            });
          }
        })
        .catch((error) => {
          console.log(error);
        });
    },
    Update_data_glue_side() {
      this.result_glue_side1 = 0;
      this.result_glue_side2 = 0;
      this.mu_seq = 1515;

      const params = new FormData();
      for (var i = 0; i < this.row_glue_side.length; i++) {
        params.append("MU_SEQ", this.re[0].mu_seq);
        params.append("GLUE_SIDE1", this.row_glue_side[i].glue_side1);
        params.append("GLUE_SIDE2", this.row_glue_side[i].glue_side2);
        params.append("GLUE_NO", this.row_glue_side[i].glue_no);
        
        axios({
          method: "post",
          url: this.$api_url + "/somain.php/update_data_glue_side",
          data: params,
        })
          .then((resp) => {
          
          })
          .catch((error) => {
            console.log(error);
          });
        this.result_glue_side1 =
          Number(this.result_glue_side1) +
          Number(this.row_glue_side[i].glue_side1);

        this.result_glue_side2 =
          Number(this.result_glue_side2) +
          Number(this.row_glue_side[i].glue_side2);
      }
      this.result_glue_side1 = Number(this.result_glue_side1) / 3;
      this.result_glue_side2 = Number(this.result_glue_side2) / 3;
      this.rowmain[0].avgglueside =
        Number(this.result_glue_side1) + Number(this.result_glue_side2);
      this.save_status_glue_side = true;
      this.status_glue = false;
    },
    opensplice() {
      this.status_splice = true;
    },
    Save_data_splice() {
      var count = 0;
      for (var i = 0; i < this.row_splice.length; i++) {
        if (this.row_splice[i].splice_value !== "") {
          count++;
        }
      }
      // console.log(count)
      this.result_splice = 0;
      this.mu_seq = 153;
      for (var i = 0; i < this.row_splice.length; i++) {
        if (this.row_splice[i].splice_value !== "") {
          const params = new FormData();
          params.append("MU_SEQ", this.re[0].mu_seq);
          params.append("splice_value", this.row_splice[i].splice_value);
          axios
            .get(this.$api_url + "/somain.php/create_splice_no")
            .then((resp) => {
              this.splice_no = resp.data.data[0].NEXTVAL;

              params.append("splice_no", this.splice_no);

              
              void axios({
                method: "post",
                url: this.$api_url + "/somain.php/save_data_splice",
                data: params,
              }).then((resp) => {
                console.log(resp.data);
              });
            });
          this.result_splice =
            Number(this.result_splice) +
            Number(this.row_splice[i].splice_value);
        }
      }
      this.last_result_splice = Number(this.result_splice) / Number(count);
      this.rowmain[0].spliceinch = this.last_result_splice;
      this.save_status_splice = true;
      this.status_splice = false;
    },
    check_splice() {
      this.row_splice = [];
      this.status_splice = true;
      this.mu_seq = 153;

      const params2 = new FormData();
      params2.append("mu_seq", this.re[0].mu_seq);
      axios({
        method: "post",
        url: this.$api_url + "/somain.php/check_splice",
        data: params2,
      })
        .then((resp) => {
          
          if (resp.data.data.length > 0) {
            resp.data.data.forEach((e) => {
              this.row_splice.push({
                splice_value: e.SPLICE,
                splice_no: e.SPLICE_NO,
              }),
                console.log(this.row_splice);
            });
          }
        })
        .catch((error) => {
          console.log(error);
        });
    },
    Update_data_splice() {
      var count = 0;
      for (var i = 0; i < this.row_splice.length; i++) {
        if (this.row_splice[i].splice_value !== "") {
          count++;
        }
      }
      // console.log(count)
      this.result_splice = 0;
      this.mu_seq = 153;
      for (var i = 0; i < this.row_splice.length; i++) {
        if (this.row_splice[i].splice_value !== "") {
          if (
            this.row_splice[i].splice_no == "" ||
            this.row_splice[i].splice_no == undefined
          ) {
            const params = new FormData();
            params.append("MU_SEQ", this.re[0].mu_seq);
            params.append("splice_value", this.row_splice[i].splice_value);

            axios
              .get(this.$api_url + "/somain.php/create_splice_no")
              .then((resp) => {
                this.splice_no = resp.data.data[0].NEXTVAL;

                params.append("splice_no", this.splice_no);

                for (var pair of params.entries()) {
                  console.log(pair[0] + ", " + pair[1]);
                }
                void axios({
                  method: "post",
                  url: this.$api_url + "/somain.php/save_data_splice",
                  data: params,
                }).then((resp) => {
                  console.log(resp.data);
                });
              });
          } else {
            const params = new FormData();
            params.append("MU_SEQ", this.re[0].mu_seq);
            params.append("SPLICE_VALUE", this.row_splice[i].splice_value);
            params.append("SPLICE_NO", this.row_splice[i].splice_no);

            for (var pair of params.entries()) {
              console.log(pair[0] + ", " + pair[1]);
            }
            void axios({
              method: "post",
              url: this.$api_url + "/somain.php/update_splice",
              data: params,
            }).then((resp) => {
              console.log(resp.data);
            });
          }
        }
        this.result_splice =
          Number(this.result_splice) + Number(this.row_splice[i].splice_value);
      }

      this.last_result_splice = Number(this.result_splice) / Number(count);
      this.rowmain[0].spliceinch = this.last_result_splice;
      this.status_splice = false;
    },
    calculate_issue(index) {
      if(this.rowwidth[index].issue_yard.length > 0){
     
      this.rowwidth[index].issuex =
        (this.rowwidth[index].issue_yard * this.rowx[0].GM_YD) / 1000;
  
      this.rowwidth[index].issue_kg = parseFloat(
        this.rowwidth[index].issuex
      ).toFixed(2);
     
      }
    },
    calavg() {
      if (this.endloss_auto !== "" && this.endloss_auto_inch !== "") {
       
        if (
          this.row_glue_side[0].glue_side1 == "" &&
          this.row_glue_side[0].glue_side2 == ""
        ) {
          this.result1 =
            Number(this.endloss_auto * 36) + Number(this.endloss_auto_inch);
       

          this.result2 =
            Number(this.rowx[0].LENGTH_YD * 36) +
            Number(this.rowx[0].LENGTH_INCH);
      
          this.result =
            ((Number(this.result1) - Number(this.result2)) *
              this.rowmain[0].ply) /
            36;
 
          this.rowmain[0].avgglueside = this.result;
        }
      } else {
        if (this.rowmain[0].avgglueside == "") {
          this.rowmain[0].avgglueside = "";
        }
      }
    },
    open_avg() {
      this.status_avg_end = true;
      if (this.avg_end_code == "EA") {
        this.avg_end_code = "EA ตั้งเผื่อเกินมาตรฐาน";
      }
      if (this.avg_end_code == "EB") {
        this.avg_end_code = "EB เผื่อ End เกินมาตรฐาน";
      }
      if (this.avg_end_code == "EC") {
        this.avg_end_code = "EC วางผ้าไม่สม่ำเสมอ";
      }
      if (this.avg_end_code == "ED") {
        this.avg_end_code =
          "ED  ปัญหาผ้า เช่น ผ้าม้วน ผ้าริมตึงกลางหย่อน ตรงกลางโค้ง";
      }
      if (this.avg_end_code == "EE") {
        this.avg_end_code =
          "EE ผิดพลาดจากพนักงาน เช่น ไม่ใช้อุปกรณ์ช่วย, พนักงานเข้าใจผิด, ใช้มาร์คผิด";
      }
      if (this.avg_end_code == "EF") {
        this.avg_end_code =
          "EF ผิดพลาดจากอุปกรณ์ เครื่องจักร เช่น มาร์คสั้นหรือยาวกว่าความจริง เครื่องปูผ้าตั้งความยาวไม่ตรงความจริง";
      }
      if (this.avg_end_code == "EG") {
        this.avg_end_code =
          "EG ใช้มาร์คสั้นกว่ามาตราฐานที่กำหนด";
      }
    },
    Save_data_avg_end() {
      this.mu_seq = 151;
      const params = new FormData();
      if (this.avg_end_code == "EA ตั้งเผื่อเกินมาตรฐาน") {
        this.avg_end_code = "EA";
        this.avg_end_cause = "ตั้งเผื่อเกินมาตรฐาน";
      }
      if (this.avg_end_code == "EB เผื่อ End เกินมาตรฐาน") {
        this.avg_end_code = "EB";
        this.avg_end_cause = "เผื่อ End เกินมาตรฐาน";
      }
      if (this.avg_end_code == "EC วางผ้าไม่สม่ำเสมอ") {
        this.avg_end_code = "EC";
        this.avg_end_cause = "วางผ้าไม่สม่ำเสมอ";
      }
      if (
        this.avg_end_code ==
        "ED ปัญหาผ้า เช่น ผ้าม้วน ผ้าริมตึงกลางหย่อน ตรงกลางโค้ง"
      ) {
        this.avg_end_code = "ED";
        this.avg_end_cause =
          "ปัญหาผ้า เช่น ผ้าม้วน ผ้าริมตึงกลางหย่อน ตรงกลางโค้ง";
      }
      if (
        this.avg_end_code ==
        "EE ผิดพลาดจากพนักงาน เช่น ไม่ใช้อุปกรณ์ช่วย, พนักงานเข้าใจผิด, ใช้มาร์คผิด"
      ) {
        this.avg_end_code = "EE";
        this.avg_end_cause =
          "ผิดพลาดจากพนักงาน เช่น ไม่ใช้อุปกรณ์ช่วย, พนักงานเข้าใจผิด, ใช้มาร์คผิด";
      }
      if (
        this.avg_end_code ==
        "EF ผิดพลาดจากอุปกรณ์ เครื่องจักร เช่น มาร์คสั้นหรือยาวกว่าความจริง เครื่องปูผ้าตั้งความยาวไม่ตรงความจริง"
      ) {
        this.avg_end_code = "EF";
        this.avg_end_cause =
          "ผิดพลาดจากอุปกรณ์ เครื่องจักร เช่น มาร์คสั้นหรือยาวกว่าความจริง เครื่องปูผ้าตั้งความยาวไม่ตรงความจริง";
      }

      if (
        this.avg_end_code ==
        "EG ใช้มาร์คสั้นกว่ามาตราฐานที่กำหนด"
      ) {
        this.avg_end_code = "EG";
        this.avg_end_cause =
          "ใช้มาร์คสั้นกว่ามาตราฐานที่กำหนด";
      }
      params.append("AVG_END_CODE", this.avg_end_code);
      params.append("AVG_END_CAUSE", this.avg_end_cause);
      params.append("MU_SEQ", this.re[0].mu_seq);
      params.append("SO_NO_DOC", this.posts.SO_NO_DOC);
      let username = this.$q.localStorage.getItem("username");
      params.append("username", username);

   

      axios({
        method: "post",
        url: this.$api_url + "/somain.php/save_avg_end_code",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
        })
        .catch((error) => {
          console.log(error);
        });
    },
    open_splice_loss_per() {
      this.status_splice_loss_per = true;
      if (this.per_splice_loss_code == "SA") {
        this.per_splice_loss_code = "SA ทำการต่อเกยจำนวนมาก";
      }
      if (this.per_splice_loss_code == "SB") {
        this.per_splice_loss_code = "SB วาดรอยต่อเกยเกินจากระยะจริง";
      }
      if (this.per_splice_loss_code == "SC") {
        this.per_splice_loss_code =
          "SC พนักงานปูทำการต่อเกยเกินจากระยะที่กำหนด";
      }
      if (this.per_splice_loss_code == "SD") {
        this.per_splice_loss_code = "SD พนักงานปูกำหนดรอยต่อเกยเอง";
      }
    },
    Save_splice_loss_per() {
      const params = new FormData();
      if (this.per_splice_loss_code == "SA  ทำการต่อเกยจำนวนมาก") {
        this.per_splice_loss_code = "SA";
        this.per_splice_loss_cause = "ทำการต่อเกยจำนวนมาก";
      }
      if (this.per_splice_loss_code == "SB วาดรอยต่อเกยเกินจากระยะจริง") {
        this.per_splice_loss_code = "SB";
        this.per_splice_loss_cause = "วาดรอยต่อเกยเกินจากระยะจริง";
      }
      if (
        this.per_splice_loss_code ==
        "SC พนักงานปูทำการต่อเกยเกินจากระยะที่กำหนด"
      ) {
        this.per_splice_loss_code = "SC";
        this.per_splice_loss_cause = "พนักงานปูทำการต่อเกยเกินจากระยะที่กำหนด";
      }
      if (this.per_splice_loss_code == "SD พนักงานปูกำหนดรอยต่อเกยเอง") {
        this.per_splice_loss_code = "SD";
        this.per_splice_loss_cause = "พนักงานปูกำหนดรอยต่อเกยเอง";
      }
      params.append("SPLICE_LOSS_CODE", this.per_splice_loss_code);
      params.append("SPLICE_LOSS_CAUSE", this.per_splice_loss_cause);
      params.append("MU_SEQ", this.re[0].mu_seq);
      params.append("SO_NO_DOC", this.posts.SO_NO_DOC);
      let username = this.$q.localStorage.getItem("username");
      params.append("username", username);

     
      axios({
        method: "post",
        url: this.$api_url + "/somain.php/save_splice_loss_per",
        data: params,
      })
        .then((resp) => {
        
        })
        .catch((error) => {
          console.log(error);
        });
    },
    open_total_cut_out_per() {
      this.status_total_cut_out_per = true;
      if (this.total_cut_out_per_code == "CA") {
        this.total_cut_out_per_code = "CA ตัดผ้าตำหนิจำนวนมาก";
      }
      if (this.total_cut_out_per_code == "CB") {
        this.total_cut_out_per_code =
          "CB พนักงานตัดผ้าตำหนิไม่เป็นไปตามมาตรฐานที่กำหนด";
      }
    },
    Save_total_cut_out_per() {
      const params = new FormData();
      if (this.total_cut_out_per_code == "CA ตัดผ้าตำหนิจำนวนมาก") {
        this.total_cut_out_per_code = "CA";
        this.total_cut_out_per_cause = "ตัดผ้าตำหนิจำนวนมาก";
      }
      if (
        this.total_cut_out_per_code ==
        "CB พนักงานตัดผ้าตำหนิไม่เป็นไปตามมาตรฐานที่กำหนด"
      ) {
        this.total_cut_out_per_code = "CB";
        this.total_cut_out_per_cause =
          "พนักงานตัดผ้าตำหนิไม่เป็นไปตามมาตรฐานที่กำหนด";
      }

      params.append("TOTAL_CUT_OUT_PER_CODE", this.total_cut_out_per_code);
      params.append("TOTAL_CUT_OUT_PER_CAUSE", this.total_cut_out_per_cause);
      params.append("MU_SEQ", this.re[0].mu_seq);
      params.append("SO_NO_DOC", this.posts.SO_NO_DOC);
      let username = this.$q.localStorage.getItem("username");
      params.append("username", username);

      for (var pair of params.entries()) {
        console.log(pair[0] + ", " + pair[1]);
      }
      axios({
        method: "post",
        url: this.$api_url + "/somain.php/save_total_cut_out_per",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
        })
        .catch((error) => {
          console.log(error);
        });
    },

    open_remnants_loss() {
      this.status_remnants_loss = true;
      if (this.remnants_loss_code == "RA") {
        this.remnants_loss_code =
          "RA พนักงานตัดหัว/ท้ายผ้าไม่เป็นไปตามมาตรฐานที่กำหนด";
      }
      if (this.remnants_loss_code == "RB") {
        this.remnants_loss_code =
          "RB ปัญหาผ้า เช่น มีตำหนิกว้างมาก หัวผ้าเอียง";
      }
    },
    Save_remnants_loss() {
      const params = new FormData();
      if (this.remnants_loss_code == "RA") {
        this.remnants_loss_code =
          "RA ปัญหาผ้า เช่น มีตำหนิกว้างมาก หัวผ้าเอียง";
        this.remnants_loss_cause =
          "พนักงานตัดหัว/ท้ายผ้าไม่เป็นไปตามมาตรฐานที่กำหนด";
      }
      if (
        this.remnants_loss_code ==
        "RB ปัญหาผ้า เช่น มีตำหนิกว้างมาก หัวผ้าเอียง"
      ) {
        this.remnants_loss_code = "RB";
        this.remnants_loss_cause = "ปัญหาผ้า เช่น มีตำหนิกว้างมาก หัวผ้าเอียง";
      }

      params.append("REMNANTS_LOSS_CODE", this.remnants_loss_code);
      params.append("REMNANTS_LOSS_CAUSE", this.remnants_loss_cause);
      params.append("MU_SEQ", this.re[0].mu_seq);
      params.append("SO_NO_DOC", this.posts.SO_NO_DOC);
      let username = this.$q.localStorage.getItem("username");
      params.append("username", username);

      for (var pair of params.entries()) {
        console.log(pair[0] + ", " + pair[1]);
      }
      axios({
        method: "post",
        url: this.$api_url + "/somain.php/save_remnants_loss",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
        })
        .catch((error) => {
          console.log(error);
        });
    },
    open_widthloss() {
      this.status_widthloss = true;
      if (this.widthloss_code == "A") {
        this.widthloss_code = "A แก้ไขมาร์คไม่มีผล";
      }
      if (this.widthloss_code == "B") {
        this.widthloss_code =
          "B ไม่ได้คัดแยกหน้าผ้าปู หน้าผ้าไม่สม่ำเสมอต่างพับ";
      }
      if (this.widthloss_code == "C") {
        this.widthloss_code = "C หน้าผ้าไม่สม่ำเสมอในพับเดียวกัน";
      }

      if (this.widthloss_code == "D ") {
        this.widthloss_code = "D ใช้มาร์คไม่เหมาะสมกับหน้าผ้า";
      }

      if (this.widthloss_code == "E") {
        this.widthloss_code =
          "E ปัญหาผ้า เช่น ริมเสียบางช่วง, ริมย้วย, ริมม้วน, รูเข็มไม่เท่ากัน";
      }
      if (this.widthloss_code == "F") {
        this.widthloss_code = "ปูผ้าหลายสีหน้าผ้าไม่เท่ากัน";
      }
    },

    Save_widthloss() {
      const params = new FormData();
      if (this.widthloss_code == "A แก้ไขมาร์คไม่มีผล") {
        this.widthloss_code = "A";
        this.widthloss_cause = "แก้ไขมาร์คไม่มีผล";
      }
      if (
        this.widthloss_code ==
        "B ไม่ได้คัดแยกหน้าผ้าปู หน้าผ้าไม่สม่ำเสมอต่างพับ"
      ) {
        this.widthloss_code = "B";
        this.widthloss_cause =
          "ไม่ได้คัดแยกหน้าผ้าปู หน้าผ้าไม่สม่ำเสมอต่างพับ";
      }
      if (this.widthloss_code == "C หน้าผ้าไม่สม่ำเสมอในพับเดียวกัน") {
        this.widthloss_code = "C";
        this.widthloss_cause = "หน้าผ้าไม่สม่ำเสมอในพับเดียวกัน";
      }

      if (this.widthloss_code == "D ใช้มาร์คไม่เหมาะสมกับหน้าผ้า") {
        this.widthloss_code = "D";
        this.widthloss_cause = "ใช้มาร์คไม่เหมาะสมกับหน้าผ้า";
      }

      if (
        this.widthloss_code ==
        "E ปัญหาผ้า เช่น ริมเสียบางช่วง, ริมย้วย, ริมม้วน, รูเข็มไม่เท่ากัน"
      ) {
        this.widthloss_code = "E";
        this.widthloss_cause =
          "ปัญหาผ้า เช่น ริมเสียบางช่วง, ริมย้วย, ริมม้วน, รูเข็มไม่เท่ากัน";
      }
      if (this.widthloss_code == "F ปูผ้าหลายสีหน้าผ้าไม่เท่ากัน") {
        this.widthloss_code = "F";
        this.widthloss_cause = "ปูผ้าหลายสีหน้าผ้าไม่เท่ากัน";
      }

      params.append("WIDTHLOSS_CODE", this.widthloss_code);
      params.append("WIDTHLOSS_CAUSE", this.widthloss_cause);
      params.append("MU_SEQ", this.re[0].mu_seq);
      params.append("SO_NO_DOC", this.posts.SO_NO_DOC);
      let username = this.$q.localStorage.getItem("username");
      params.append("username", username);

      for (var pair of params.entries()) {
        console.log(pair[0] + ", " + pair[1]);
      }
      axios({
        method: "post",
        url: this.$api_url + "/somain.php/save_widthloss",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
        })
        .catch((error) => {
          console.log(error);
        });
    },
   async confirm_delete(){
      this.count_row_gm--
      this.index = this.delete_index

  

      const params = new FormData();
      params.append("mu_seq", this.re[0].mu_seq);
      params.append("gm_no", this.rowx[this.index].gm_no);

 await    axios({
        method: "post",
        url: this.$api_url + "/somain.php/delete_detail",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
          if (resp.data.status == true) {
            this.$q.notify({
              message:
                "DELETE Successful" +
                " " +
                "GM NO:" +
                this.exrow +
                " " +
                "MU SEQ:" +
                this.re[0].mu_seq,
              color: "green-9",
            });
          }
        })
        .catch((error) => {
          console.log(error);
        });

  await    axios({
        method: "post",
        url: this.$api_url + "/somain.php/delete_detail_sub",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
        })
        .catch((error) => {
          console.log(error);
        });

      const params3 = new FormData();
          this.sum_value_ply = 0
          params3.append("MU_SEQ", this.re[0].mu_seq)
  await   axios({
          method: "post",
          url: this.$api_url + "/somain.php/keep_data_for_ply",
          data: params3,
        }).then((resp)=>{
          console.log(resp.data)
          this.rowexport=[]
           if (resp.data.data.length > 0) {
            resp.data.data.forEach((e) => {
              this.rowexport.push({
                mu_seq:e.MU_SEQ,
                gm_no:e.GM_NO,
                ply:e.PLY
              });
            });
        for(var ax =0; ax<this.rowexport.length; ax++){
          console.log(this.rowexport[ax].ply)
        this.sum_value_ply = Number(this.sum_value_ply) + Number(this.rowexport[ax].ply)
        }
        console.log(this.sum_value_ply)
       
        
        
         const params5 = new FormData();
        params5.append("SO_NO", this.SO_NO)
        params5.append("SO_YEAR", this.SO_YEAR)
        params5.append("GM_NO", this.rowexport[0].gm_no)
        params5.append("PLY", this.sum_value_ply);

     
        

         axios({
          method: "post",
          url: this.$api_url + "/somain.php/check_qty_cut_2",
          data: params5,
        })
          .then((resp) => {
            console.log(resp.data);
            this.data_top_qty_cut = resp.data.data[0];
          })
          .catch((error) => {
            console.log(error);
          });

    
          }
        })  
      this.rowx.splice(this.index, 1);

      this.check_bottom_data();
       this.confirm_delete_detail = false
    },
    cancel_delete(){
      this.confirm_delete_detail = false
    },

    Back() {
      window.history.back();
    },
  },
  mounted() {
    let org = this.$q.localStorage.getItem("org");
    this.org = org;
    this.start_date = new Date();
    this.start= moment(this.start_date).format('DD/MM/YYYY')

  },
  watch: {
    cpartno: function (value) {
      this.checkgm(value);
    },
/*      start(val){
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

        }  */
  },
};
</script>
<style lang="sass">
.my-card
  padding-top: 20px
  padding-right: 10px
  padding-bottom: 20px
  padding-left: 10px
.my-sticky


  thead tr th
    position: sticky
    z-index: 1
  thead tr:first-child th
    top: 0

.my-stickys


  thead tr:nth-child(12) th:nth-child(12),
  td:nth-child(12)
    position: sticky
    -webkit-position: sticky /* add this line */
    right: 0
    z-index: 1

  thead tr:nth-child(11) th:nth-child(11),
  td:nth-child(11)
    position: sticky
    -webkit-position: sticky /* add this line */
    right: 0
    z-index: 1
</style>
