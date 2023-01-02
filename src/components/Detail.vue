<template>
  <center>
    <q-btn color="secondary"  label="Detail" @click="check_Data()" />
    <q-dialog v-model="prompt" persistent>
      <q-card style="min-width: 100%; min-height: 100%">
        <q-banner
          expand="(true)"
          dense
          inline-actions
          class="text-white bg-indigo-9"
        >
          Data: {{ this.posts.SO_NO }}
          <template v-slot:action>
            <q-btn icon="cancel" flat color="white" v-close-popup />
          </template>
        </q-banner>
        <q-form @submit="search" class="q-gutter-md">
          <div class="row">
            <div class="col-1"></div>
            <div class="col-2">
              <q-input
                input-class="text-center"
                v-model="posts.ORG"
                :dense="dense"
                label="ORG"
                disable
              />
            </div>
            <div class="col-1"></div>
            <div class="col-2">
              <q-input
                input-class="text-center"
                v-model="posts.SO_YEAR"
                :dense="dense"
                label="Year"
                disable
              />
            </div>
            <div class="col-2"></div>
            <div class="col-2">
              <q-input
                input-class="text-center"
                v-model="posts.SO_NO"
                :dense="dense"
                label="SO"
                disable
              />
            </div>
            <div class="col-2"></div>
          </div>
        </q-form>
        <div class="row">
          <div class="col-1"></div>
          <div class="col-2">
            <q-input
              v-model="posts.SO_NO_DOC"
              :dense="dense"
              label="SO_NO_DOC"
              disable
            />
          </div>
          <div class="col-1"></div>
          <div class="col-3">
            <q-input
              v-model="posts.STYLE_REF"
              :dense="dense"
              label="STYLE"
              disable
            />
          </div>
          <div class="col-1"></div>
          <div class="col-3">
            <q-input
              v-model="posts.CUSTOMER_NAME"
              :dense="dense"
              label="CUSTOMER"
              disable
            />
          </div>
        </div>
        <div class="row">
          <div class="col-1"></div>
          <div class="col-2">
            <q-input
              v-model="posts.ORDER_QTY"
              :dense="dense"
              label="QTY ORDER"
              disable
            />
          </div>

          <div class="col-1"></div>
          <div class="col-3">
            <q-input
              v-model="data_top_qty_cut2.CUT_QTY"
              :dense="dense"
              label="QTY CUT"
              disable
            />
          </div>
          <div class="col-1"></div>
          <div class="col-3">
            <q-input
              v-model="posts.GMT_TYPE"
              :dense="dense"
              label="GMT TYPE"
              disable
            />
          </div>
          <div class="col-1"></div>
        </div>
        <div class="row">
          <div class="col-1"></div>
          <div class="col-2">
            <q-input
              v-model="posts.F_TYPE"
              :dense="dense"
              label="F TYPE"
              disable
            />
          </div>
          <div class="col-1"></div>
          <div class="col-7">
            <q-input
              v-model="posts.FABRIC_TYPE"
              :dense="dense"
              label="Fabric TYPE"
              disable
            />
          </div>

          <div class="col-1"></div>
        </div>
        <br />
        <br />
        <div class="row">
          <div class="q-px-md"></div>
          <div class="col-1">
            <q-btn icon="add" size="md" color="green-9" @click="addRowgm" />
          </div>

          <div class="col-2">
            <q-input filled v-model="posts.MU_DATE">
              <template v-slot:append>
                <q-icon name="event" class="cursor-pointer">
                  <q-popup-proxy
                    ref="qDateProxy"
                    transition-show="scale"
                    transition-hide="scale"
                  >
                    <q-date v-model="posts.MU_DATE" mask="DD/MM/YYYY">
                      <div class="row items-center justify-end">
                        <q-btn
                          v-close-popup
                          label="Close"
                          color="primary"
                          flat
                        />
                      </div>
                    </q-date>
                  </q-popup-proxy>
                </q-icon>
              </template>
            </q-input>
          </div>
          <div class="col-1"></div>
          <div class="col-1">
            <q-input
              outlined
              v-model="posts.CPART_NO"
              desne
              disable
              label="Cpart NO"
            />
          </div>
          <div class="col-1"></div>
          <div class="col-1">
            <q-input
              outlined
              v-model="endloss_auto.END_LOSS_AUTO"
              label="Endloss auto"
            />
          </div>
          <div class="col-1"></div>
          <div class="col-1">
            <q-input
              outlined
              v-model="endloss_auto.END_LOSS_INCH"
              label="Endloss auto Inch"
            />
          </div>
          <div class="col-1"></div>
          <div class="col-1">
            <q-btn
              icon="delete"
              size="xl"
              color="negative"
              @click="delete_All"
            ></q-btn>
          </div>
        </div>

        <br />
        <br />
        <div class="row">
          <div class="col-12">
            <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
              <q-table
                :rows="row_show_detail_gm"
                :columns="columns_show_detail_gm"
                hide-pagination
                table-header-style="text-align:center;"
                table-header-class="bg-blue-4 text-black"
                :rows-per-page-options="[10]"
                row-key="GM_NO"
                dense
                style="min-height: 100px"
              >
                <template v-slot:body="props">
                  <q-tr :props="props">
                    <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                    <!--   -->

                    <q-td key="name" :props="props">
                      <!-- {{props.row.gm}}
                                                                    <select v-model="props.row.gm"  style="min-width: 100px;" @change="gmselect(props.row.gm)">
                                                                        <option v-for="item in gm" value="" :key="item.name" >--Select--</option>
                                                                        <option v-for="item in gm" value={{item.name}} :key="item.name" >{{item.name}}</option>
                                                                    </select> -->
                    </q-td>
                    <q-td key="number" :props="props">
                      <q-input
                        input-class="text-center"
                        v-model="props.row.MU_SEQ"
                        disable
                        label=""
                        size="4"
                      />
                    </q-td>
                    <q-td key="number" :props="props" style="min-width: 120px">
                      <q-select
                        v-if="props.row.status == false"
                        v-model="props.row.GM_NO"
                        :options="option_gm_select"
                        @update:model-value="gmselect(props.pageIndex)"
                        label="SELECT GM"
                      />
                      <q-input
                        v-else
                        input-class="text-center"
                        v-model="props.row.GM_NO_DISPLAY"
                        disable
                        label=""
                      />
                    </q-td>
                    <q-td key="number" :props="props">
                    <q-select outlined v-model="props.row.TABLE" :options="chose_table" label="Table" />
                    </q-td>
                    <q-td key="weightm2" :props="props">
                      <q-input
                        input-class="text-center"
                        v-model="props.row.WEIGHTM2"
                        label="weightm2"
                      />
                    </q-td>
                    <q-td key="weightyd" :props="props">
                      <q-input
                        type="number"
                        input-class="text-center"
                        v-model="props.row.WEIGHTYD"
                        label="weightyd"
                      />
                    </q-td>
                    <q-td key="widthinch" :props="props">
                      <q-input
                        type="number"
                        input-class="text-center"
                        v-model="props.row.WIDTH_INCH"
                        label="widthinch"
                      />
                    </q-td>
                    <q-td key="lenthyds" :props="props">
                      <q-input
                        disable
                        input-class="text-center"
                        v-model="props.row.YD"
                        label="lenthyds"
                      />
                    </q-td>
                    <q-td key="lenthyd" :props="props">
                      <q-input
                        type="number"
                        input-class="text-center"
                        v-model="props.row.LENGTH_YD"
                        label="lenthyd"
                      />
                    </q-td>

                    <q-td key="lenthinch" :props="props">
                      <q-input
                        input-class="text-center"
                        v-model="props.row.LENGTH_INCH"
                        label="lenthinch"
                      />
                    </q-td>
                    <q-td key="ttlmarker" :props="props">
                      <q-input
                        disable
                        input-class="text-center"
                        v-model="props.row.TTL_MARKER"
                        label="ttlmarker"
                      />
                    </q-td>
                    <q-td key="marker" :props="props">
                      <q-input
                        input-class="text-center"
                        v-model="props.row.MARKER"
                        label="marker"
                      />
                    </q-td>
                    <q-td key="ply" :props="props">
                      <q-input
                        input-class="text-center"
                        v-model="props.row.PLY"
                        label="ply"
                      />
                    </q-td>
                    <q-td key="ply" :props="props">
                      <q-btn
                        icon="add"
                        color="green-9"
                        @click="updateDetail(props.pageIndex)"
                      ></q-btn>
                    </q-td>
                    <q-td key="ply" :props="props">
                      <q-btn
                        icon="save"
                        color="primary"
                        @click="updateData(props.pageIndex)"
                      ></q-btn>
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
        <br />
        <q-card style="min-width: 90%">
          <q-card-section class="q-pt-none">
            <q-tabs v-model="tab" dense align="justify" narrow-indicator>
              <q-tab
                name="main"
                class="bg-blue-10 text-white"
                label="MAIN"
                @click="calavg()"
              />

              <q-tab
                name="result"
                class="bg-blue-10 text-white"
                label="Result"
                @click="calavg()"
              />
            </q-tabs>
          </q-card-section>
          <q-tab-panels v-model="tab" animated>
            <q-tab-panel name="main">
              <div class="row">
                <div class="col-12">
                  <div
                    class="q-pt-md"
                    style="margin-top: 0px; padding-top: 0px"
                  >
                    <q-table
                      :rows="rowmain"
                      :columns="comlumns_bottom_1"
                      table-header-class="bg-blue-3 text-black"
                      :rows-per-page-options="[10]"
                      hide-pagination
                      dense
                      row-key="cause"
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
                              :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                              :disable="isDisabled2"
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
                               :bg-color="isDisabled2 ? 'blue-grey-1' : ''"
                              :disable="isDisabled2"
                              dense
                              label=""
                            />
                          </q-td>
                          <q-td key="length_inch" :props="props">
                            <q-input
                              input-class="text-center"
                              type="number"
                              v-model="props.row.length_inch"
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
                  <div
                    class="q-pt-md"
                    style="margin-top: 0px; padding-top: 0px"
                  >
                    <q-table
                      :rows="rowmain"
                      :columns="comlumns_bottom_2"
                      :rows-per-page-options="[10]"
                      table-header-class="bg-blue-3 text-black"
                      hide-pagination
                      row-key="lenghth_yd"
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
                              input-class="text-center"
                              type="number"
                              v-model="props.row.spliceinch"
                              dense
                              @click="opensplice()"
                              label=""
                            />
                          </q-td>
                                <q-td key="spliceinch" :props="props">
                       <q-btn
                        push
                          align="left"
                          color="primary"
                    
                        
                          @click="opensplice()"
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
                  <div
                    class="q-pt-md"
                    style="margin-top: 0px; padding-top: 0px"
                  >
                    <q-table
                      :rows="rowmain"
                      :columns="comlumns_bottom_3"
                      :rows-per-page-options="[10]"
                      table-header-class="bg-blue-3 text-black"
                      row-key="cutoutnoof"
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
                  <div
                    class="q-pt-md"
                    style="margin-top: 0px; padding-top: 0px"
                  >
                    <q-table
                      :rows="rowmain"
                      :columns="comlumns_bottom_4"
                      :rows-per-page-options="[10]"
                      row-key="endloss_yd"
                      table-header-class="bg-blue-3 text-black"
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
                              input-class="text-center"
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
                         
                        
                            @click="open_glue_side()"
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
                              v-if="
                                this.rowmain[0].avg_end > this.fix_avg_end_loss
                              "
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
                                this.rowmain[0].splicelossper >
                                this.fix_splice_loss
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
                  <div
                    class="q-pt-md"
                    style="margin-top: 0px; padding-top: 0px"
                  >
                    <q-table
                      :rows="rowmain"
                      :columns="comlumns_bottom_5"
                      :rows-per-page-options="[10]"
                      table-header-class="bg-blue-3 text-black"
                      row-key="splicelossper"
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
                                this.rowmain[0].rement_per >
                                this.fix_remnants_loss
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
        </q-card>

        <div class="q-pa-md q-gutter-sm">
          <q-dialog
            v-model="basic"
            transition-show="rotate"
            transition-hide="rotate"
          >
            <q-card style="min-width: 50%">
              <q-card-section>
                <div class="row">
                  <div class="col-3">
                    <div class="text-h6">Detail data</div>
                  </div>
                  <div class="col-8"></div>
                  <div class="col-1">
                    <div>
                      <q-btn icon="cancel" flat color="black" v-close-popup />
                    </div>
                  </div>
                </div>

                <q-card-actions align="right">
                  <q-btn
                    icon="add"
                    color="positive"
                    size="md"
                    @click="addRow"
                  />
                  <q-btn
                    icon="save"
                    size="md"
                    color="primary"
                    @click="Updateselect"
                  />
                </q-card-actions>
              </q-card-section>

              <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
                <q-table
                  :rows="rowwidth"
                  :columns="columns3"
                  class="my-sticky"
                  table-header-class="bg-blue-4 text-black"
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
                      <q-td key="issue_yard" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.issue_yard"
                          @update:model-value="calculate_issue(props.pageIndex)"
                          label=""
                        />
                      </q-td>
                      <q-td key="issue_kg" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.issue_kg"
                          label=""
                        />
                      </q-td>

                      <q-td key="avg_width" :props="props">
                        <q-input
                          input-class="text-center"
                          type="number"
                          v-model="props.row.avg_width"
                          label=""
                        />
                      </q-td>
                      <q-td key="delete" :props="props">
                        <!--key = name.columne -->
                        <q-btn
                          color="red-12"
                          icon="delete"
                          @click="deleteSelected(props.pageIndex)"
                        />
                      </q-td>
                    </q-tr>
                  </template>
                </q-table>
              </div>
            </q-card>
          </q-dialog>
          <q-card-actions align="center">
            <q-btn color="indigo-10" label="Update" @click="Update_Bottom()" />
          </q-card-actions>
        </div>
      </q-card>
    </q-dialog>
    <div class="q-pa-md q-gutter-sm">
      <q-dialog
        v-model="check_delete"
        transition-show="rotate"
        transition-hide="rotate"
      >
        <q-card style="min-width: 30%">
          <q-card-section>
            <div class="text-h6">Confirm Delete</div>
            <q-card-actions align="right">
              <q-btn
                color="green-4"
                size="lg"
                icon="done"
                @click="confirm_delete()"
              />
              <q-btn
                color="negative"
                size="lg"
                icon="cancel"
                @click="cancel_delete()"
              />
            </q-card-actions>
          </q-card-section>
        </q-card>
      </q-dialog>
    </div>

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
              color="positive"
              size="md"
              @click="addRowsplice()"
            />

            <q-btn
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
      v-model="status_glue"
      transition-show="rotate"
      transition-hide="rotate"
    >
      <q-card style="min-width: 50%">
        <q-card-section>
          <div class="text-h6">Insert ข้อมูล Avg Glue Side</div>
          <q-card-actions align="right">
            <q-btn
              icon="add"
              color="positive"
              size="md"
              @click="addRow_glue_side()"
            />
            <q-btn
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

        <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
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

        <q-dialog
          v-model="confirm_delete_detail"
          transition-show="rotate"
          transition-hide="rotate"
        >
          <q-card style="min-width: 30%">
            <q-card-section>
              <div class="text-h6">Confirm Delete</div>
              <q-card-actions align="right">
                <q-btn color="green-4" size="lg" icon="done" @click="confirm_delete_2()"/>
                <q-btn color="negative" size="lg" icon="cancel"  @click="cancel_delete_2()" />
              </q-card-actions>
            </q-card-section>

           
          </q-card>
        </q-dialog>

    
  </center>
  
</template>

<script>
import axios from "axios";
import { ref } from "vue";
import { useQuasar, QSpinnerGears } from "quasar";
import { onBeforeUnmount } from "vue";
export default {
  setup() {
    const $q = useQuasar();
    let timer;
    const $api_url = "http://172.16.160.119:8080/sosystem/api";

    onBeforeUnmount(() => {
      if (timer !== void 0) {
        clearTimeout(timer);
        this.$q.loading.hide();
      }
    });
    return {
      prompt: ref(false),
      basic: ref(false),
      tab: ref("main"), //อย่าลืมใส่
    };
  },
  props: {
    detail: {
      type: String,
    },
    detail3: {
      type: String,
    },
    detail4: {
      type: String,
    },
    detail5: {
      type: String,
    },
    detail6: {
      type: String,
    },
  },
  watch: {
    detail: function (val) {
      if (val) {
        this.detail2 = val;
        console.log(this.detail2);
      }
    },
  },
  data() {
    return {
      code: ["A", "B", "C", "D", "E"],
      option_gm_select: [],
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
      count_row_gm_select: 0,
      status: true,
      rowexport:[],
      status_update: true,
      gm_data:ref(""),
      check_status: false,
      data_line_no:"",
      confirm_delete_detail:ref(false),
      delete_index:'',
      result_qty_cut:'',
      keep_layer:[],
      chose_table: ["A1", "A2", "A3", "A4","A5","A6","A7","A8","A9","A10","A11","A12",
      "A13","A14","A15","A16","A17","A18","A19","A20","M1","M2","M3","M4","M5","M6","M7","M8","M9","M10"],

      rowy: [],
      date: "",
      columns_show_detail_gm: [
        {
          name: "MU_SEQ",
          label: "MU_SEQ",
          align: "center",
          field: "MU_SEQ",
        },
        {
          name: "gm_no",
          label: "GM_NO",
          align: "center",
          field: "GM_NO",
        },
        {
          name: "number",
          label: "Table",
          align: "center",
          field: "TABLE",
        },
        {
          name: "weightm2",
          label: "GM/M2",
          align: "center",
          field: "weightm2",
        },
        {
          name: "weightyd",
          label: "GM/YD",
          align: "center",
          field: "weightyd",
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
          name: "detail",
          label: "Detail",
          align: "center",
          field: "detail",
        },
        {
          name: "update",
          label: "Update",
          align: "center",
          field: "update",
        },
        {
          name: "delete",
          label: "Delete",
          align: "center",
          field: "delete",
        },
      ],

      rowmain: [],
      columns3: [
        {
          name: "no",
          label: "No.",
          align: "center",
          field: "no",
        },

        {
          name: "issue_yard",
          label: "Issue YARD",
          align: "center",
          field: "issue_yard",
        },
        {
          name: "issue_kg",
          label: "Issue KG",
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
      rowwidth: [],
      rowshow: [],

      rowyx: [
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

      posts: [],
      gm_no: [],

      rowz: [],
      rowxy: [],
      line_no: [],
      row_show_detail_gm: [],
      status1: false,
      status2: false,
      status3: false,
      isDisabled2: true,
      check_delete: false,
      status_splice: false,
      status_glue: false,
      data_top_qty_cut:[],
      data_top_qty_cut2:[],
      endloss_auto: [],
      endloss_inch: [],
    };
  },

  methods: {
    addRow() {
      this.rowwidth.push([
        {
          no: "",
          issue_kg: "",
          issue_yard: "",
          avg_width: "",
          delete: "",
        },
      ]);
    },
    deleteSelected(index) {
     

      const params = new FormData();
      params.append("MU_SEQ", this.detail6);
      params.append("GM_NO", this.gm_data);
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
    check_Data() {
      this.$q.loading.show({
        spinner: QSpinnerGears,
        spinnerColor: "wthite",
        spinnerSize: 180,
        backgroundColor: "black",
      });

      this.prompt = true;

      const params = new FormData();
      params.append("SO_NO", this.detail);
      params.append("SO_YEAR", this.detail3);
      params.append("SO_NO_DOC", this.detail4);
      params.append("CPART", this.detail5);
      params.append("MU_SEQ", this.detail6);

      axios({
        method: "post",
        url: this.$api_url + "/detail.php/checkso",
        data: params,
      }).then((resp) => {
        this.posts = resp.data.data[0];
      });

      axios({
        method: "post",
        url: this.$api_url + "/detail.php/check_endless_auto",
        data: params,
      }).then((resp) => {
        console.log(resp.data);
        this.endloss_auto = resp.data.data[0];
        console.log(this.endloss_auto);
      });

      if (this.status1 == false) {
        axios({
          method: "post",
          url: this.$api_url + "/detail.php/checkgmno",
          data: params,
        }).then((resp) => {
          console.log(resp.data);
          if (resp.data.data.length > 0) {
            resp.data.data.forEach((e) => {
              this.row_show_detail_gm.push({
                MU_SEQ: e.MU_SEQ,
                GM_NO: e.GM_NO,
                GM_NO_DISPLAY: e.GM_NO.substring(7),
                TABLE: e.CUT_TABLE,
                WEIGHTM2: e.WEIGHT_M2,
                WEIGHTYD: e.WEIGHT_YD,
                WIDTH_INCH: e.WIDTH_INCH,
                LENGTH_YD: e.LENGTH_YD,
                LENGTH_INCH: e.LENGTH_INCH,
                YD: parseFloat(
                  Number(e.LENGTH_INCH / 36) + Number(e.LENGTH_YD)
                ).toFixed(2),
                TTL_MARKER: (
                  (parseFloat(e.LENGTH_INCH) / 36 + parseFloat(e.LENGTH_YD)) *
                  parseFloat(e.PLY)
                ).toFixed(2),
                MARKER: e.MARKER,
                PLY: e.PLY,
              });
            });
          }
          
        });
        const params3 = new FormData();
          this.sum_value_ply = 0
          params3.append("MU_SEQ", this.detail6)
         axios({
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
        params5.append("SO_NO", this.detail)
        params5.append("SO_YEAR", this.detail3)
        params5.append("GM_NO", this.rowexport[0].gm_no)
        params5.append("PLY", this.sum_value_ply);

     
        

         axios({
          method: "post",
          url: this.$api_url + "/somain.php/check_qty_cut_2",
          data: params5,
        })
          .then((resp) => {
            console.log(resp.data);
            this.data_top_qty_cut2 = resp.data.data[0];
          })
          .catch((error) => {
            console.log(error);
          });

    
          }
        })
         
        this.status1 = true;
        console.log(this.status1);
      }

      if (this.status2 == false) {
        this.check_bottom_data();

        this.status2 = true;
        console.log(this.status2);
      }

      // hiding in 3s
      this.timer = setTimeout(() => {
        this.$q.loading.hide();
        this.timer = void 0;
      }, 800);
    },

    chosegm() {
      const params = new FormData();

      params.append("GM_NO", this.GM_NO);

      axios({
        method: "post",
        url: this.$api_url + "/detail.php/findrowgm",
        data: params,
      }).then((resp) => {
        this.row_show_detail_gm = resp.data.data;
        console.log(this.row_show_detail_gm[0].GM_NO);
      });
    },
    async updateData(index) {
      
      this.data_top_qty_cut2 = []
      this.rowexport=[]

      if(this.row_show_detail_gm[index].TABLE == '' || this.row_show_detail_gm[index].TABLE == undefined){
          this.$q.notify({
                          message: "Please Insert Table",
                          position: "center",
                          color: "red-9",
                        });
      }else{
      const params = new FormData();
      params.append("MU_SEQ", this.row_show_detail_gm[index].MU_SEQ);
      params.append("GM_NO", this.row_show_detail_gm[index].GM_NO);
      params.append("CUT_TABLE", this.row_show_detail_gm[index].TABLE);
      params.append("WEIGHT_M2", this.row_show_detail_gm[index].WEIGHTM2);
      params.append("WEIGHT_YD", this.row_show_detail_gm[index].WEIGHTYD);
      params.append("WIDTH_INCH", this.row_show_detail_gm[index].WIDTH_INCH);
      params.append("LENGTH_YD", this.row_show_detail_gm[index].LENGTH_YD);
      this.row_show_detail_gm[index].YD = parseFloat(
        Number(this.row_show_detail_gm[index].LENGTH_INCH / 36) +
          Number(this.row_show_detail_gm[index].LENGTH_YD)
      ).toFixed(2);
      params.append("LENGTH_INCH", this.row_show_detail_gm[index].LENGTH_INCH);

      params.append("MARKER", this.row_show_detail_gm[index].MARKER);
      params.append("PLY", this.row_show_detail_gm[index].PLY);
      let username = this.$q.localStorage.getItem("username");
      params.append("username", username);
      params.append("CPART_NO", this.posts.CPART_NO);

      for (var pair of params.entries()) {
        console.log(pair[0] + ", " + pair[1]);
      }
     

     axios({
        method: "post",
        url: this.$api_url + "/detail.php/check_data_have",
        data: params,
      }).then((resp) => {
        console.log(resp.data.data.length)
        if (resp.data.data.length == 0) {
          axios({
            method: "post",
            url: this.$api_url + "/detail.php/create_detail",
            data: params,
          })
            .then((resp) => {
              console.log(resp.data);
             this.row_show_detail_gm=[];
             axios({
          method: "post",
          url: this.$api_url + "/detail.php/checkgmno",
          data: params,
        }).then((resp) => {
          console.log(resp.data);
          if (resp.data.data.length > 0) {
            resp.data.data.forEach((e) => {
              this.row_show_detail_gm.push({
                MU_SEQ: e.MU_SEQ,
                GM_NO: e.GM_NO,
                GM_NO_DISPLAY: e.GM_NO.substring(7),
                TABLE: e.CUT_TABLE,
                WEIGHTM2: e.WEIGHT_M2,
                WEIGHTYD: e.WEIGHT_YD,
                WIDTH_INCH: e.WIDTH_INCH,
                LENGTH_YD: e.LENGTH_YD,
                LENGTH_INCH: e.LENGTH_INCH,
                YD: parseFloat(
                  Number(e.LENGTH_INCH / 36) + Number(e.LENGTH_YD)
                ).toFixed(2),
                TTL_MARKER: (
                  (parseFloat(e.LENGTH_INCH) / 36 + parseFloat(e.LENGTH_YD)) *
                  parseFloat(e.PLY)
                ).toFixed(2),
                MARKER: e.MARKER,
                PLY: e.PLY,
              });
            });
             this.$q.notify({
                          message: "Update Detail Success",
                          position: "center",
                          color: "green-9",
                        });
          }
          
           });
            })
            .catch((error) => {
              console.log(error);
            });
             
    const params3 = new FormData();
    params3.append("MU_SEQ", this.row_show_detail_gm[index].MU_SEQ)
     axios({
          method: "post",
          url: this.$api_url + "/somain.php/keep_data_for_ply",
          data: params3,
        }).then((resp)=>{
           this.rowexport=[]
          console.log(resp.data)
           if (resp.data.data.length > 0) {
            resp.data.data.forEach((e) => {
              this.rowexport.push({
                mu_seq:e.MU_SEQ,
                gm_no:e.GM_NO,
                ply:e.PLY
              });
            });
            for(var ax =0; ax<this.rowexport.length; ax++){
          
        this.sum_value_ply = Number(this.sum_value_ply) + Number(this.rowexport[ax].ply)
        }
          }
        })


        
        console.log(this.sum_value_ply)
         const params5 = new FormData();
        params5.append("SO_NO", this.detail)
        params5.append("SO_YEAR", this.detail3)
        params5.append("GM_NO", this.row_show_detail_gm[index].GM_NO)
        params5.append("PLY", this.sum_value_ply);

     
        

      axios({
          method: "post",
          url: this.$api_url + "/somain.php/check_qty_cut_2",
          data: params5,
        })
          .then((resp) => {
            console.log(resp.data);
            this.data_top_qty_cut2 = resp.data.data[0];
          })
          .catch((error) => {
            console.log(error);
          });

         

        
        }else{
                 
            axios({
                    method: "post",
                    url: this.$api_url + "/detail.php/updatedata",
                    data: params,
                  })
                    .then((resp) => {
                      console.log(resp.data);
                      if (resp.data.status == false) {
                        this.$q.notify({
                          message: "Update Detail Success",
                          position: "center",
                          color: "green-9",
                        });
                       
                      }
                    })
                    .catch((error) => {
                      cosole.log(error);
                    }); 

        
          const params3 = new FormData();
          this.sum_value_ply = 0
          params3.append("MU_SEQ", this.row_show_detail_gm[index].MU_SEQ)
         axios({
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
        params5.append("SO_NO", this.detail)
        params5.append("SO_YEAR", this.detail3)
        params5.append("GM_NO", this.row_show_detail_gm[index].GM_NO)
        params5.append("PLY", this.sum_value_ply);

     
        

         axios({
          method: "post",
          url: this.$api_url + "/somain.php/check_qty_cut_2",
          data: params5,
        })
          .then((resp) => {
            console.log(resp.data);
            this.data_top_qty_cut2 = resp.data.data[0];
          })
          .catch((error) => {
            console.log(error);
          });

    
          }
        })
        


        
                } 
      });
       this.rowmain = [];
      this.check_bottom_data();
      }
    },
    updateDetail(index) {
      
      
     
      this.rowwidth = [];
      console.log(this.rowwidth);
      console.log(this.detail6);
      console.log(this.row_show_detail_gm[index].GM_NO);

      const params = new FormData();
      params.append("MU_SEQ", this.detail6);
      params.append("GM_NO", this.row_show_detail_gm[index].GM_NO);
      
      axios({
         method: "post",
        url: this.$api_url + "/detail.php/check_data_have",
        data: params,
      }).then((resp)=>{
        console.log(resp.data)
        console.log(resp.data.data.length )
        if(resp.data.data.length > 0){
           this.basic = true;
           axios({
        method: "post",
        url: this.$api_url + "/detail.php/show_sub_detail",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data)
          if (resp.data.data.length > 0) {
            resp.data.data.forEach((e) => {
              this.rowwidth.push({
                no: e.LINE_NO,
                issue_kg: e.ISSUE,
                issue_yard: e.ISSUE_YRD,
                avg_width: e.AVG_WIDTH,
                gm_no: e.GM_NO,
                delete: "",
              }),
                console.log(this.rowwidth);
            });
          }

         
        })
        .catch((error) => {
          console.log(error);
        });
        }else{
          this.$q.notify({
                message: "Please Save GM",
                position: "center",
                color: "red-9",
              });
        }
      })
     
        this.gm_data = this.row_show_detail_gm[index].GM_NO
    },
    change_cause() {
      console.log(this.rowmain[0].code);
      if (this.rowmain[0].code == "A") {
        this.rowmain[0].cause = "มาร์คไม่มีผล";
        console.log(this.rowmain[0].cause);
      }
      if (this.rowmain[0].code == "B") {
        this.rowmain[0].cause = "ไม่ได้คัดแยกหน้าผ้าปู";
        console.log(this.rowmain[0].cause);
      }
      if (this.rowmain[0].code == "C") {
        this.rowmain[0].cause = "หน้าผ้าไม่สม่ำเสมอในม้วน และต่างม้วน";
        console.log(this.rowmain[0].cause);
      }
      if (this.rowmain[0].code == "D") {
        this.rowmain[0].cause = "ใช้มาร์คไม่เหมาะสมกับหน้าผ้า";
        console.log(this.rowmain[0].cause);
      }
      if (this.rowmain[0].code == "E") {
        this.rowmain[0].cause =
          "อื่น ๆ ริมเสียบางช่วง, ริมย้วย, ปูผ้าหลายสีหน้าผ้าไม่เท่ากัน";
        console.log(this.rowmain[0].cause);
      }
    },
    Updateselect(index) {
       for (var i = 0; i < this.rowwidth.length; i++) {
        this.issue2 = this.rowwidth[i].issue_kg;
        this.issue3 = this.rowwidth[i].issue_yard;
        this.avg_width2 = this.rowwidth[i].avg_width;


        if (this.rowwidth[i].no !== undefined) {
          const params2 = new FormData();
          params2.append("MU_SEQ2", this.detail6);
          params2.append("GM_NO2", this.gm_data);
          if (this.issue2 == undefined || this.issue2 == null) {
            this.issue2 = "";
          }
          params2.append("ISSUE2", this.issue2);
          if (this.issue3 == undefined || this.issue3 == null) {
            this.issue3 = "";
          }
          params2.append("ISSUE2_YARD", this.issue3);
          if (this.avg_width2 == undefined || this.avg_width2 == null) {
            this.avg_width2 = "";
          }
          params2.append("AVG_WIDTH2", this.avg_width2);
          params2.append("LINE_NO", this.rowwidth[i].no);
           for (var pair of params2.entries()) {
          console.log(pair[0] + ", " + pair[1]);
      }
          
          axios({
            method: "post",
            url: this.$api_url + "/detail.php/update_sub_detail",
            data: params2,
          })
            .then((resp) => {
              console.log(resp.data)
              this.$q.notify({
                message: "Update Detail Success",
                position: "center",
                color: "green-9",
              });
             
           
            })
            .catch((error) => {
              console.log(error);
            });
            
     
        }else{
        
           const params2 = new FormData();
          params2.append("MU_SEQ2", this.detail6);
          params2.append("GM_NO2", this.gm_data);
         
          if (this.issue3 == undefined || this.issue3 == null) {
            this.issue3 = "";
          }
          params2.append("ISSUE2_YARD", this.issue3);
           if (this.issue2 == undefined || this.issue2 == null) {
            this.issue2 = "";
          }
          params2.append("ISSUE2", this.issue2);
          if (this.avg_width2 == undefined || this.avg_width2 == null) {
            this.avg_width2 = "";
          }
          params2.append("AVG_WIDTH2", this.avg_width2);

          axios
            .get(this.$api_url + "/somain.php/create_line_no")
            .then((resp) => {
              this.line_no = resp.data.data[0].NEXTVAL;
            
              params2.append("LINE_NO", this.line_no);
          axios({
                method: "post",
                url: this.$api_url + "/detail.php/insert_sub_detail",
                data: params2,
              })
                .then((resp) => {
                  console.log(resp.data);
                 
                })
            }); 
       
        }
      }
      this.rowmain = [];  
      this.check_bottom_data();
      this.basic = false;
      
    
    },
    Update_Bottom() {
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
      if (this.rowmain[0].splicenoof == undefined) {
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
        this.rowmain[0].spliceinch == undefined ||
        this.rowmain[0].spliceinch == "NaN"
      ) {
        this.rowmain[0].spliceinch = "";
      }
      if (
        this.endloss_auto.END_LOSS_AUTO == null ||
        this.endloss_auto.END_LOSS_AUTO == undefined
      ) {
        this.endloss_auto.END_LOSS_AUTO = "";
      }
      if (
        this.endloss_auto.END_LOSS_INCH == null ||
        this.endloss_auto.END_LOSS_INCH == undefined
      ) {
        this.endloss_auto.END_LOSS_INCH = "";
      }
      if (
        this.rowmain[0].re_cut == "NaN" ||
        this.rowmain[0].re_cut == undefined
      ) {
        this.rowmain[0].re_cut = "";
      }
       if (
        this.rowmain[0].endplug12_yd == "NaN" ||
        this.rowmain[0].endplug12_yd == undefined
      ) {
        this.rowmain[0].endplug12_yd = "";
      }
        if (
        this.rowmain[0].plug12_inch == "NaN" ||
        this.rowmain[0].plug12_inch == undefined
      ) {
        this.rowmain[0].plug12_inch = "";
      }
         if (
        this.rowmain[0].plug12 == "NaN" ||
        this.rowmain[0].plug12 == undefined
      ) {
        this.rowmain[0].plug12 = "";
      }


      const params = new FormData();
      params.append("MU_SEQ", this.row_show_detail_gm[0].MU_SEQ);
      params.append("SO_NO_DOC", this.posts.SO_NO_DOC);
      params.append("SO_NO", this.posts.SO_NO);

      params.append("endloss_auto", this.endloss_auto.END_LOSS_AUTO);
      params.append("end_plug1", this.rowmain[0].endplug1);
      params.append("end_plug2", this.rowmain[0].endplug2);
      params.append("casue", this.rowmain[0].cause);
      params.append("code", this.rowmain[0].code);
      params.append("splicenoof", this.rowmain[0].splicenoof);
      params.append("spliceinch", this.rowmain[0].spliceinch);
      params.append("cutoutnoof", 0);
      params.append("cutoutweight", this.rowmain[0].cutoutweight);
      params.append("remnants", this.rowmain[0].remnants);

      params.append("endloss_inch", this.endloss_auto.END_LOSS_INCH);
      params.append("widthplug1", this.rowmain[0].widthplug1);
      params.append("widthplug2", this.rowmain[0].widthplug2);
      params.append("avgglueside", this.rowmain[0].avgglueside);
      params.append("re_cut", this.rowmain[0].re_cut);
      params.append("endplug12_yd", this.rowmain[0].endplug12_yd);
      params.append("plug12_inch", this.rowmain[0].plug12_inch);
      params.append("plug12", this.rowmain[0].plug12);


      for (var pair of params.entries()) {
        console.log(pair[0] + ", " + pair[1]);
      }

      axios({
        method: "post",
        url: this.$api_url + "/detail.php/update_bottom",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
          if (resp.data.status == false) {
            this.$q.notify({
              message: "Update Success",

              color: "green-9",
            });
            this.rowmain = [];
            this.check_bottom_data();
          }
        })
        .catch((error) => {
          console.log(error);
        });
    },
    check_bottom_data(){
      this.rowmain = [];
       const params = new FormData();
      params.append("SO_NO", this.detail);
      params.append("SO_YEAR", this.detail3);
      params.append("SO_NO_DOC", this.detail4);
      params.append("CPART", this.detail5);
      params.append("MU_SEQ", this.detail6);
      axios({
        //all detail ด้านล่าง
        method: "post",
        url: this.$api_url + "/detail.php/all_data_bottom",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
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
    delete_All() {
      this.check_delete = true;
    },
    confirm_delete() {
      
      this.rowshow = [];
      const params = new FormData();
      params.append("SO_NO", this.detail);
      params.append("mu_seq", this.detail6);
      params.append("so_no_doc", this.detail4);
      params.append("SO_NO_DOC", this.detail4);

      axios({
        method: "post",
        url: this.$api_url + "/detail.php/delete_detail_sub",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
        })
        .catch((error) => {
          console.log(error);
        });

      axios({
        method: "post",
        url: this.$api_url + "/detail.php/delete_detail",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
        })
        .catch((error) => {
          console.log(error);
        });

      axios({
        method: "post",
        url: this.$api_url + "/detail.php/delete_head",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
          if (resp.data.status == true) {
          }
        })
        .catch((error) => {
          console.log(error);
        });
      axios({
        method: "post",
        url: this.$api_url + "/detail.php/delete_glue_side",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
          if (resp.data.status == true) {
          }
        })
        .catch((error) => {
          console.log(error);
        });
      axios({
        method: "post",
        url: this.$api_url + "/detail.php/delete_splice",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
          if (resp.data.status == true) {
          }
        })
        .catch((error) => {
          console.log(error);
        });
      this.$q.notify({
        message: "Delete MU_SEQ :" + this.detail6 + " " + "Sucess",
        position: "center",
        color: "red-9",
      });
      this.rowshow = [];
      axios({
        method: "post",
        url: this.$api_url + "/mainso.php/checksearch",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
          console.log(this.rowshow);
          resp.data.data.forEach((e) => {
            this.rowshow.push({
              MU_SEQ: e.MU_SEQ,
              SO_NO: e.SO_NO,
              SO_NO_DOC: e.SO_NO_DOC,
              SO_YEAR: e.SO_YEAR,
              GM_NO: e.GM_NO,
              MU_DATE: e.MU_DATE,
              CPART: e.CPART_NO,
              STYLE: e.STYLE_REF,
              CUST: e.CUSTOMER_NAME,
            });
          });
          this.$emit("show", this.rowshow);
        })
        .catch((error) => {
          console.log(error);
        });
      this.check_delete = false;
      this.prompt = false;
    },
    opensplice() {
      this.status_splice = true;
      this.row_splice = [];

      const params2 = new FormData();
      params2.append("mu_seq", this.detail6);
      axios({
        method: "post",
        url: this.$api_url + "/somain.php/check_splice",
        data: params2,
      })
        .then((resp) => {
          console.log(resp.data);
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
    addRow_glue_side() {
      if (this.row_glue_side.length > 2) {
        this.$q.notify({
          message: "Can't Add row more than 3 rows",

          color: "green-9",
        });
      } else {
        this.row_glue_side.push([
          {
            no: "",
            glue_side1: "",
            glue_side2: "",
            delete: "",
          },
        ]);
      }
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
            params.append("MU_SEQ", this.detail6);
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
            params.append("MU_SEQ", this.detail6);
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
      console.log(this.rowwidth[index].issue_yard);
      console.log(this.row_show_detail_gm[0].WEIGHTYD);
      this.rowwidth[index].issuex =
        (this.rowwidth[index].issue_yard *
          this.row_show_detail_gm[0].WEIGHTYD) /
        1000;
      console.log(this.rowwidth[index].issuex);
      this.rowwidth[index].issue_kg = parseFloat(
        this.rowwidth[index].issuex
      ).toFixed(2);
      console.log(this.rowwidth[index].issue_kg);
      }
    },
    open_glue_side() {
      this.status_glue = true;
      this.row_glue_side = [];
      this.status_glue = true;

      const params2 = new FormData();
      params2.append("mu_seq", this.detail6);
      axios({
        method: "post",
        url: this.$api_url + "/somain.php/check_data_glue_side",
        data: params2,
      })
        .then((resp) => {
          console.log(resp.data);
          if (resp.data.data.length > 0) {
            resp.data.data.forEach((e) => {
              this.row_glue_side.push({
                glue_no: e.GLUE_NO,
                glue_side1: e.GLUE_SIDE1,
                glue_side2: e.GLUE_SIDE2,
              });
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

      for (var i = 0; i < this.row_glue_side.length; i++) {
        if (this.row_glue_side[i].glue_no == undefined) {
          const params = new FormData();
          params.append("MU_SEQ", this.detail6);
          params.append("GLUE_SIDE1", this.row_glue_side[i].glue_side1);
          params.append("GLUE_SIDE2", this.row_glue_side[i].glue_side2);

          for (var pair of params.entries()) {
            console.log(pair[0] + ", " + pair[1]);
          }

          axios
            .get(this.$api_url + "/somain.php/createmu_glue_side")
            .then((resp) => {
              this.line_no = resp.data.data[0].NEXTVAL;

              params.append("glue_no", this.line_no);
              for (var pair of params.entries()) {
                console.log(pair[0] + ", " + pair[1]);
              }

              void axios({
                method: "post",
                url: this.$api_url + "/somain.php/save_data_glue_side",
                data: params,
              }).then((resp) => {
                console.log(resp.data);
              });
            });
        } else {
          const params = new FormData();
          params.append("MU_SEQ", this.detail6);
          params.append("GLUE_SIDE1", this.row_glue_side[i].glue_side1);
          params.append("GLUE_SIDE2", this.row_glue_side[i].glue_side2);
          params.append("GLUE_NO", this.row_glue_side[i].glue_no);
          axios({
            method: "post",
            url: this.$api_url + "/somain.php/update_data_glue_side",
            data: params,
          })
            .then((resp) => {
              console.log(resp.data);
            })
            .catch((error) => {
              console.log(error);
            });
        }
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
    open_avg() {
      const params = new FormData();
      params.append("MU_SEQ", this.detail6);
      params.append("SO_NO_DOC", this.detail4);
      axios({
        method: "post",
        url: this.$api_url + "/detail.php/open_avg",
        data: params,
      }).then((resp) => {
        this.avg_end_code = resp.data.data[0].AVG_END_CODE;
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
      });
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
        this.avg_end_code == "EG ใช้มาร์คสั้นกว่ามาตราฐานที่กำหนด"
      ) {
        this.avg_end_code = "EG";
        this.avg_end_cause ="ใช้มาร์คสั้นกว่ามาตราฐานที่กำหนด"
      }
      params.append("AVG_END_CODE", this.avg_end_code);
      params.append("AVG_END_CAUSE", this.avg_end_cause);
      params.append("MU_SEQ", this.detail6);
      params.append("SO_NO_DOC", this.posts.SO_NO_DOC);
      let username = this.$q.localStorage.getItem("username");
      params.append("username", username);

      for (var pair of params.entries()) {
        console.log(pair[0] + ", " + pair[1]);
      }

      axios({
        method: "post",
        url: this.$api_url + "/somain.php/save_avg_end_code",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
          this.$q.notify({
            message: "Update Success",

            color: "green-9",
          });
        })
        .catch((error) => {
          console.log(error);
        });
    },
    open_splice_loss_per() {
      const params = new FormData();
      params.append("MU_SEQ", this.detail6);
      params.append("SO_NO_DOC", this.detail4);
      axios({
        method: "post",
        url: this.$api_url + "/detail.php/open_avg",
        data: params,
      }).then((resp) => {
        console.log(resp.data);
        console.log(resp.data.data[0].SPLICE_LOSS_PER_CODE);
        this.per_splice_loss_code = resp.data.data[0].SPLICE_LOSS_PER_CODE;
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
      });
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
      params.append("MU_SEQ", this.detail6);
      params.append("SO_NO_DOC", this.posts.SO_NO_DOC);
      let username = this.$q.localStorage.getItem("username");
      params.append("username", username);

      for (var pair of params.entries()) {
        console.log(pair[0] + ", " + pair[1]);
      }

      axios({
        method: "post",
        url: this.$api_url + "/somain.php/save_splice_loss_per",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
          this.$q.notify({
            message: "Update Success",

            color: "green-9",
          });
        })
        .catch((error) => {
          console.log(error);
        });
    },
    open_total_cut_out_per() {
      const params = new FormData();
      params.append("MU_SEQ", this.detail6);
      params.append("SO_NO_DOC", this.detail4);
      axios({
        method: "post",
        url: this.$api_url + "/detail.php/open_avg",
        data: params,
      }).then((resp) => {
        this.total_cut_out_per_code = resp.data.data[0].TOTAL_CUT_OUT_PER_CODE;
        this.status_total_cut_out_per = true;
        if (this.total_cut_out_per_code == "CA") {
          this.total_cut_out_per_code = "CA ตัดผ้าตำหนิจำนวนมาก";
        }
        if (this.total_cut_out_per_code == "CB") {
          this.total_cut_out_per_code =
            "CB พนักงานตัดผ้าตำหนิไม่เป็นไปตามมาตรฐานที่กำหนด";
        }
      });
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
      params.append("MU_SEQ", this.detail6);
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
          this.$q.notify({
            message: "Update Success",

            color: "green-9",
          });
        })
        .catch((error) => {
          console.log(error);
        });
    },

    open_remnants_loss() {
      const params = new FormData();
      params.append("MU_SEQ", this.detail6);
      params.append("SO_NO_DOC", this.detail4);
      axios({
        method: "post",
        url: this.$api_url + "/detail.php/open_avg",
        data: params,
      }).then((resp) => {
        this.remnants_loss_code = resp.data.data[0].REMNANTS_LOSS_CODE;
        this.status_remnants_loss = true;
        if (this.remnants_loss_code == "RA") {
          this.remnants_loss_code =
            "RA พนักงานตัดหัว/ท้ายผ้าไม่เป็นไปตามมาตรฐานที่กำหนด";
        }
        if (this.remnants_loss_code == "RB") {
          this.remnants_loss_code =
            "RB ปัญหาผ้า เช่น มีตำหนิกว้างมาก หัวผ้าเอียง";
        }
      });
    },
    Save_remnants_loss() {
      const params = new FormData();
      if (
        this.remnants_loss_code ==
        "RA พนักงานตัดหัว/ท้ายผ้าไม่เป็นไปตามมาตรฐานที่กำหนด"
      ) {
        this.remnants_loss_code = "RA";
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
      params.append("MU_SEQ", this.detail6);
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
          this.$q.notify({
            message: "Update Success",

            color: "green-9",
          });
        })
        .catch((error) => {
          console.log(error);
        });
    },
    open_widthloss() {
      const params = new FormData();
      params.append("MU_SEQ", this.detail6);
      params.append("SO_NO_DOC", this.detail4);
      axios({
        method: "post",
        url: this.$api_url + "/detail.php/open_avg",
        data: params,
      }).then((resp) => {
        this.widthloss_code = resp.data.data[0].WIDTH_LOSS_CODE;
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
          this.widthloss_code = "F ปูผ้าหลายสีหน้าผ้าไม่เท่ากัน";
        }
      });
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
      params.append("MU_SEQ", this.detail6);
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
          this.$q.notify({
            message: "Update Success",

            color: "green-9",
          });
        })
        .catch((error) => {
          console.log(error);
        });
    },
    calavg() {
      if (
        this.endloss_auto.END_LOSS_AUTO !== null &&
        this.endloss_auto.END_LOSS_INCH !== null
      ) {
        this.rowmain[0].avgglueside = "";
        this.result1 =
          Number(this.endloss_auto.END_LOSS_AUTO * 36) +
          Number(this.endloss_auto.END_LOSS_INCH);

        console.log(this.result1);

        this.result2 =
          Number(this.row_show_detail_gm[0].LENGTH_YD * 36) +
          Number(this.row_show_detail_gm[0].LENGTH_INCH);
        console.log(this.result2);

        if (this.result1 == 0 || this.result2 == 0) {
          this.rowmain[0].avgglueside = "";
        } else {
          this.result =
            ((Number(this.result1) - Number(this.result2)) *
              this.rowmain[0].ply) /
            36;
          console.log(this.result);
          this.rowmain[0].avgglueside = this.result;
        }
      } else {
        if (this.rowmain[0].avgglueside == "") {
          this.rowmain[0].avgglueside = "";
        }
      }
    },
    addRowgm() {
      this.row_show_detail_gm.push({
        MU_SEQ: this.row_show_detail_gm[0].MU_SEQ,
        GM_NO: "",
        TABLE: "",
        WEIGHTM2: "",
        WEIGHTYD: "",
        WIDTH_INCH: "",
        LENGTH_YD: "",
        LENGTH_INCH: "",
        YD: "",
        TTL_MARKER: "",
        MARKER: "",
        PLY: "",
        status: false,
      });

      const params = new FormData();
      params.append("SO_YEAR", this.detail3);
      params.append("MU_SEQ",this.detail6);
      params.append("SO_NO", this.detail);
      params.append("cpartno", this.posts.CPART_NO);
      this.rowxy = [];
      axios({
        method: "post",
        url: this.$api_url + "/somain.php/checkgm",
        data: params,
      }).then((resp) => {
        console.log(resp.data);

        var k = 0;

        resp.data.data.forEach((e) => {
          this.option_gm_select[k] = e.GM_NO; //นำค่าเข้าสู่ rowz
          k++;
        });
        console.log(this.option_gm_select);
      });
    },

    deleteRow(index) {
      this.confirm_delete_detail = true
      this.delete_index = index
      
    },
    gmselect(index) {
      const params = new FormData();
      params.append("GM_NO", this.row_show_detail_gm[index].GM_NO);
      params.append("SO_YEAR", this.detail3);
      params.append("MU_SEQ",this.detail6);
      params.append("SO_NO", this.detail);
      params.append("CPART_NO", this.posts.CPART_NO);
      for (var pair of params.entries()) {
        console.log(pair[0] + ", " + pair[1]);
      }
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
                  for (let xx = 0; xx <= this.row_show_detail_gm.length; xx++) {
                    if (xx == index) {
                      this.row_show_detail_gm[index].gm_no = e.GM_NO;
                      this.row_show_detail_gm[index].WEIGHTM2 = e.GM_M2;
                      this.row_show_detail_gm[index].WEIGHTYD = e.GM_YD;
                      this.row_show_detail_gm[index].WIDTH_INCH = e.WIDTH_INCH;
                      this.row_show_detail_gm[index].LENGTH_YD = e.LENGTH_YD;
                      this.row_show_detail_gm[index].YD = e.YD;
                      this.row_show_detail_gm[index].LENGTH_INCH =
                        e.LENGTH_INCH;
                      this.row_show_detail_gm[index].TTL_MARKER = parseFloat(
                        e.YD * e.LAYER
                      ).toFixed(2);
                      this.row_show_detail_gm[index].MARKER = e.PER_U;
                      this.row_show_detail_gm[index].PLY = e.LAYER;
                      console.log(xx);
                      if (xx > 0) {
                        if (
                          this.row_show_detail_gm[xx].GM_M2 !==
                          this.row_show_detail_gm[0].GM_M2
                        ) {
                          this.$q.notify({
                            message: "ข้อมูล GM_M2 ไม่ตรงกัน",

                            color: "red-9",
                          });
                        }

                        if (
                          this.row_show_detail_gm[xx].GM_YD !==
                          this.row_show_detail_gm[0].GM_YD
                        ) {
                          this.$q.notify({
                            message: "ข้อมูล GM_YD ไม่ตรงกัน",

                            color: "red-9",
                          });
                        }

                        if (
                          this.row_show_detail_gm[xx].WIDTH_INCH !==
                          this.row_show_detail_gm[0].WIDTH_INCH
                        ) {
                          this.$q.notify({
                            message: "ข้อมูล Weight Inch ไม่ตรงกัน",

                            color: "red-9",
                          });
                        }

                        if (
                          this.row_show_detail_gm[xx].YD !==
                          this.row_show_detail_gm[0].YD
                        ) {
                          this.$q.notify({
                            message: "ข้อมูล Length_YD(1) ไม่ตรงกัน",

                            color: "red-9",
                          });
                        }

                        if (
                          this.row_show_detail_gm[xx].LENGTH_YD !==
                          this.row_show_detail_gm[0].LENGTH_YD
                        ) {
                          this.$q.notify({
                            message: "ข้อมูล LENGTH_YD(2) ไม่ตรงกัน",

                            color: "red-9",
                          });
                        }

                        if (
                          this.row_show_detail_gm[xx].LENGTH_INCH !==
                          this.row_show_detail_gm[0].LENGTH_INCH
                        ) {
                          this.$q.notify({
                            message: "ข้อมูล LENGTH_INCH ไม่ตรงกัน",

                            color: "red-9",
                          });
                        }

                        if (
                          this.row_show_detail_gm[xx].ttlmarker !==
                          this.row_show_detail_gm[0].ttlmarker
                        ) {
                          this.$q.notify({
                            message: "ข้อมูล ttlmarker ไม่ตรงกัน",

                            color: "red-9",
                          });
                        }

                        if (
                          this.row_show_detail_gm[xx].PER_U !==
                          this.row_show_detail_gm[0].PER_U
                        ) {
                          this.$q.notify({
                            message: "ข้อมูล Marker ไม่ตรงกัน",

                            color: "red-9",
                          });
                        }

                        if (
                          this.row_show_detail_gm[xx].LAYER !==
                          this.row_show_detail_gm[0].LAYER
                        ) {
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

      

     

      /*  if(value !== undefined)
        this.test.push(value)
        console.log(this.test)
        console.log(value) */
    },

    cancel_delete() {
      this.check_delete = false;
    },

   async confirm_delete_2(){
      this.index = this.delete_index
      console.log(this.index)
      console.log(this.detail6);
      console.log(this.row_show_detail_gm[this.index].GM_NO);

       const params = new FormData();
      params.append("mu_seq", this.detail6);
      params.append("gm_no", this.row_show_detail_gm[this.index].GM_NO);

   await   axios({
        method: "post",
        url: this.$api_url + "/somain.php/delete_detail",
        data: params,
      })
        .then((resp) => {
          console.log(resp.data);
          if (resp.data.status == true) {
            this.$q.notify({
              message:
                "DELETE Successful" ,
              color: "green-9",
            });
          }
        })
        .catch((error) => {
          console.log(error);
        });

   await   axios({
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
          params3.append("MU_SEQ", this.detail6)
         axios({
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
        params5.append("SO_NO", this.detail)
        params5.append("SO_YEAR", this.detail3)
        params5.append("GM_NO", this.rowexport[0].gm_no)
        params5.append("PLY", this.sum_value_ply);

     
        

         axios({
          method: "post",
          url: this.$api_url + "/somain.php/check_qty_cut_2",
          data: params5,
        })
          .then((resp) => {
            console.log(resp.data);
            this.data_top_qty_cut2 = resp.data.data[0];
          })
          .catch((error) => {
            console.log(error);
          });

    
          }
        })
      this.row_show_detail_gm.splice(this.index, 1);

      this.check_bottom_data();
       this.confirm_delete_detail = false 
    },
    cancel_delete_2(){
      this.confirm_delete_detail = false
    }, 
  },
  watch: {
    cpartno: function (value) {
      this.checkgm(value);
    },
  },
};
</script>

<style>
.q-banner {
  position: -webkit-sticky;
  position: sticky;
  top: 0px;

  z-index: 2;
}
</style>
