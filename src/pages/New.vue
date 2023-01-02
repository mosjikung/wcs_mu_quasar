<template>

  <div class="q-pa-md" style="max-width: 100%">
    <div class="q-gutter-md">
      <q-card class="my-card" style="max-width:100%">
      
        <div class="row">
         
          <div class="col-3">
            <q-input
              v-model="SO_YEAR"
              @keypress.enter.prevent="search()"
              type="number"
              :dense="dense"
              label="Year"
              :disable="isDisabled"
            />
          </div>
          <div class="col-1"></div>
          <div class="col-3">
            <q-input
              v-model="SO_NO"
              @keypress.enter.prevent="search()"
              type="number"
              :dense="dense"
              label="SO"
              :disable="isDisabled"
            />
          </div>
          <div class="col-1"></div>
          <div class="col-1">
            <q-btn
              v-if="this.isDisabled"
              icon="search"
              type="submit"
              label="search"
              :disable="isDisabled"
              stack
              
              
              color="blue-3 text-black"
            />
            <q-btn
              v-else
              icon="search"
              type="submit"
              label="search"
              stack
              @click="search()"
              
              
              color="primary text-white"
            />
          </div>
          <div class="col-1">
            <q-btn
              icon="edit_off"
              type="submit"
              label="clear"
              stack
              
              color="wthite text-black"
              @click="clearall()"
            />
          </div>
          <div class="col-1">
              <q-btn
              icon="arrow_back"
              type="submit"
              label="Back"
              stack
              
              color="grey-3 text-black"
              @click="Back()"
            />
        </div>
          </div>
        
    

      <div class="row">
        <div class="col-3">
          <q-input
            v-model="posts.SO_NO_DOC"
            :dense="dense"
            label="SO_NO_DOC"
            :disable="isDisabled"
          />
        </div>
        <div class="col-1"></div>
        <div class="col-3">
          <q-input
            v-model="posts.STYLE_REF"
            :dense="dense"
            label="STYLE"
            :disable="isDisabled"
          />
        </div>
        <div class="col-1"></div>
        <div class="col-3">
          <q-input
            v-model="posts.CUST_NAME"
            :dense="dense"
            label="CUSTOMER"
            :disable="isDisabled"
          />
        </div>
        <div class="col-1"></div>
      </div>

      <div class="row">
        <div class="col-3">
          <q-input
            v-model="posts.QTY"
            :dense="dense"
            label="QTY ORDER"
            :disable="isDisabled"
          />
        </div>
        <div class="col-1"></div>
        <div class="col-3">
          <q-input
            v-model="QTY_CUT"
            :dense="dense"
            label="QTY CUT"
            :disable="isDisabled"
          />
        </div>
        <div class="col-1"></div>
        <div class="col-3">
          <q-input
            v-model="posts.GMT_TYPE"
            :dense="dense"
            label="GMT TYPE"
            :disable="isDisabled"
          />
        </div>
        <div class="col-1"></div>
      </div>
      <div class="row">
        
        <div class="col-2">
         <q-select
            outlined
            v-model="F_TYPE"
            :options="f_type"
            label="F TYPE"
            :disable="isDisabled"
           
          />
     <q-tooltip v-model="Showing2">A: WOVEN - NYLON, TAFFETA.<br> B:LIGHT-MIDDLE WEIGHT - SINGLE JERSEY, INTERLOCK, DOUBLE KNIT, JACQUARD <br> C:FLEECE<br> D:SPECIAL - RIB, ELASTIC, ELASTAN, LYCRA, SPANDEX, TWO WAY, WAFFLE, PRE WASH, PRESHRUNK, VELOUR<br> </q-tooltip>
        </div>
        <div class="col-2"></div>
        <div class="col-7">
          <q-input
            v-model="posts.FABRIC_TYPE"
            :dense="dense"
            label="Fabric TYPE"
            :disable="isDisabled"
          />
        </div>
        <div class="col-1"></div>
      </div>
      </q-card>

      <div class="row">
         <div class="col-1">
          <q-btn icon="new_label" size="xl" color="green-10"  @click="addRowgm" />
        </div>
        <div class="col-1"></div>
        <div class="col-2">
          <q-input filled v-model="date" mask="date" :rules="['date']">
            <template v-slot:append>
              <q-icon name="event" class="cursor-pointer">
                <q-popup-proxy
                  ref="qDateProxy"
                  transition-show="scale"
                  transition-hide="scale"
                >
                  <q-date v-model="date">
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
        <div class="col-2">
          <q-select
            outlined
            v-model="cpartno"
            :options="rowz"
            label="Cpart"
            :disable="isDisabled"
            @onclick="checkgm()"
          />
        </div>
        <div class="col-1"></div>
        <div class="col-2">
          <q-select
            outlined
            v-model="org"
            :options="orgs"
            label="Chose Org"
            :disable="isDisabled"
          />
        </div>
        <div class="col-1"></div>
       
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
              style="min-height: 100px"
            >
            
            
            
            
              <template v-slot:body="props">
                <q-tr table-class= "bg-red-3" :props="props">
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
                    <q-input
                      input-class="text-center"
                      v-model="props.row.number"
                      label=""
                      size="4"
                    />
                  </q-td>
                  <q-td key="weightm2" :props="props">
                    <q-input
                      input-class="text-center"
                      type="number"
                      v-model="props.row.GM_M2"
                      label="weightm2"
                    />
                  </q-td>
                  <q-td key="weightyd" :props="props">
                    <q-input
                      input-class="text-center"
                      type="number"
                      v-model="props.row.GM_YD"
                      label="weightyd"
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
                  <q-td key="lenthyd" :props="props">
                    <q-input
                      input-class="text-center"
                      type="number"
                      v-model="props.row.YD"
                      label="lenthyd"
                    />
                  </q-td>
                  <q-td key="lenthyds" :props="props">
                    <q-input
                      input-class="text-center"
                      type="number"
                      v-model="props.row.LENGTH_YD"
                      label="lenthyds"
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
  <br>
  <br>
  <br>
      <q-card style="min-width: 90%">
        <q-card-section class="q-pt-none">
          <q-tabs v-model="tab" dense align="justify" narrow-indicator>
            <q-tab name="main" class="bg-blue-10 text-white" label="MAIN" />

            <q-tab name="result" class="bg-blue-10 text-white" label="Result" />
          </q-tabs>
        </q-card-section>
        <q-tab-panels v-model="tab" animated>
          <q-tab-panel name="main">
            <div class="row">
              <div class="col-12">
                <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
                  <q-table
                    :rows="rowmain"
                    :columns="comlumnsmain"
                    table-header-class="bg-blue-3 text-black"
                    :rows-per-page-options="[10]"
                    hide-pagination
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
                            label=""
                          />
                        </q-td>
                        <q-td key="gmm2" :props="props">
                          <q-input
                            
                            input-class="text-center"
                            type="number"
                            v-model="props.row.gmm2"
                            label=""
                          />
                        </q-td>
                        <q-td key="gmyd" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.gmyd"
                            label=""
                          />
                        </q-td>
                        <q-td key="lyd" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.lyd"
                            label=""
                          />
                        </q-td>
                        <q-td key="linch" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.linch"
                            label=""
                          />
                        </q-td>
                        <q-td key="obs" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.obs"
                            label=""
                          />
                        </q-td>
                        <q-td key="marker" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.marker"
                            label=""
                          />
                        </q-td>
                        <q-td key="ply" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.ply"
                            label=""
                          />
                        </q-td>
                        <q-td key="endplug1" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.endplug1"
                            label=""
                          />
                        </q-td>
                        <q-td key="endplug2" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.endplug2"
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
                    :columns="columnssec"
                    :rows-per-page-options="[10]"
                    table-header-class="bg-blue-3 text-black"
                    hide-pagination
                    row-key="cause"
                    style="min-height: 150px"
                  >
                    <template v-slot:body="props">
                      <q-tr :props="props">
                        <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                        <!--   -->

                        <q-td key="cause" :props="props">
                          <q-input
                            input-class="text-center"
                            v-model="props.row.cause"
                            label=""
                          />
                        </q-td>
                        <q-td key="code" :props="props">
                          <q-select
                            outlined
                            v-model="props.row.code"
                            :options="code"
                            @update:model-value="change_cause()"
                            label=""
                          />
                          <q-tooltip v-model="showing">A:มาร์คไม่มีผล .<br> B:ไม่ได้คัดแยกหน้าผ้าปู<br> C:หน้าผ้าไม่สม่ำเสมอในม้วน และต่างม้วน<br> D:ใช้มาร์คไม่เหมาะสมกับหน้าผ้า<br> E: อื่น ๆ ริมเสียบางช่วง, ริมย้วย, ปูผ้าหลายสีหน้าผ้าไม่เท่ากัน</q-tooltip>
                        </q-td>
                        <q-td key="splicenoof" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.splicenoof"
                            label=""
                          />
                        </q-td>
                        <q-td key="spliceinch" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.spliceinch"
                            label=""
                          />
                        </q-td>
                        <q-td key="cutoutnoof" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.cutoutnoof"
                            label=""
                          />
                        </q-td>
                        <q-td key="cutoutweight" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.cutoutweight"
                            label=""
                          />
                        </q-td>
                        <q-td key="remnants" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.remnants"
                            label=""
                          />
                        </q-td>
                        <q-td key="widthplug1" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.widthplug1"
                            label=""
                          />
                        </q-td>
                        <q-td key="widthplug2" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.widthplug2"
                            label=""
                          />
                        </q-td>
                        <q-td key="avgglueside" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.avgglueside"
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

          <q-tab-panel name="result">
            <div class="row">
              <div class="col-12">
                <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
                  <q-table
                    table-header-class="bg-blue-3 text-black"
                    :rows="rowmain"
                    :columns="columns2"
                    :rows-per-page-options="[10]"
                    row-key="yardroll"
                    hide-pagination
                    style="min-height: 150px"
                  >
                    <template v-slot:body="props">
                      <q-tr :props="props">
                        <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                        <!--   -->

                        <q-td key="yardroll" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.yard_roll"
                            label=""
                          />
                        </q-td>
                        <q-td key="lengthyd" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.lyd"
                            label=""
                          />
                        </q-td>
                        <q-td key="marker" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.obs"
                            label=""
                          />
                        </q-td>
                        <q-td key="ttlmarker" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.obswidth"
                            label=""
                          />
                        </q-td>
                        <q-td key="end" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.end"
                            label=""
                          />
                        </q-td>
                        <q-td key="plug1" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.plug1"
                            label=""
                          />
                        </q-td>
                        <q-td key="allplug1" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.perplug1"
                            label=""
                          />
                        </q-td>
                        <q-td key="plug2" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.plug2"
                            label=""
                          />
                        </q-td>
                        <q-td key="allplug2" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.perplug2"
                            label=""
                          />
                        </q-td>
                        <q-td key="plug12" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.plug12"
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
                    :columns="columnsb"
                    :rows-per-page-options="[10]"
                    row-key="sectionlosslenth"
                    hide-pagination
                    style="min-height: 150px"
                  >
                    <template v-slot:body="props">
                      <q-tr :props="props">
                        <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                        <!--   -->

                         <q-td key="endloss" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.endloss"
                            label=""
                          />
                        </q-td>
                        <q-td key="widthloss" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.widthloss"
                            label=""
                          />
                        </q-td>
                        
                        <q-td key="avglenthper" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.avglenthper"
                            label=""
                          />
                        </q-td>
                        <q-td key="avgend" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.avgend"
                            label=""
                          />
                        </q-td>
                        <q-td key="sectionlosslenth" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.sectionlosslenth"
                            label=""
                          />
                        </q-td>
                        <q-td key="sectionlossper" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.sectionlossper"
                            label=""
                          />
                        </q-td>
                        <q-td key="spliceloss" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.spliceloss"
                            label=""
                          />
                        </q-td>
                        <q-td key="splicelossper" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.splicelossper"
                            label=""
                          />
                        </q-td>
                        <q-td key="overlap" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.overlap"
                            label=""
                          />
                        </q-td>
                        <q-td key="overlapper" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.overlapper"
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
                    :columns="comlumnsc"
                    :rows-per-page-options="[10]"
                    row-key="defect"
                    hide-pagination
                    style="min-height: 150px"
                  >
                    <template v-slot:body="props">
                      <q-tr :props="props">
                        <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                        <!--   -->

                        <q-td key="defect" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.defect"
                            label=""
                          />
                        </q-td>
                        <q-td key="defectper" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.defectper"
                            label=""
                          />
                        </q-td>
                        <q-td key="totalcutout" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.totalcutout"
                            label=""
                          />
                        </q-td>
                        <q-td key="totalcutoutper" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.totalcutoutper"
                            label=""
                          />
                        </q-td>
                        <q-td key="rement" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.rement"
                            label=""
                          />
                        </q-td>
                        <q-td key="rement_per" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.rement_per"
                            label=""
                          />
                        </q-td>
                        <q-td key="totallenthloss" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.totallenthloss"
                            label=""
                          />
                        </q-td>
                        <q-td key="totallenthlossper" :props="props">
                          <q-input
                            input-class="text-center"
                            type="number"
                            v-model="props.row.totallenthlossper"
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
      <!--  <q-form @submit="Savedata()"> -->

      <div class="q-pa-md q-gutter-sm">
        <q-dialog
          v-model="basic"
          transition-show="rotate"
          transition-hide="rotate"
        >
          <q-card style="min-width: 50%">
            <q-card-section>
              <div class="text-h6">Insert ข้อมูล</div>
              <q-card-actions align="right">
                <q-btn icon="add" color="positive" size="xl"  @click="addRow" />
                <q-btn
                  icon="update"
                  size="xl"
                  color="teal-10"
                  type="submit"
                  
                  @click="Savedata()"
                  v-close-popup
                />
              </q-card-actions>
            </q-card-section>

            <div class="q-pt-md" style="margin-top: 0px; padding-top: 0px">
              <q-table
                :rows="rowwidth"
                :columns="columns3"
                class="my-sticky"
                table-header-class="bg-blue-grey-5"
                :rows-per-page-options="[20]"
                hide-pagination=""
                row-key="no"
                style="height: 380px"
              >
                <template v-slot:body="props">
                  <q-tr :props="props">
                    <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                    <!--   -->

                    <q-td key="no" :props="props">
                      {{ props.pageIndex + 1 }}
                    </q-td>
                    <q-td key="issue" :props="props">
                      <q-input
                        input-class="text-center"
                        type="number"
                        v-model="props.row.issue"
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
      </div>
      <!-- </q-form>
     -->

      <q-card-actions align="center">
        <q-btn color="indigo-10" label="Save Data" @click="saveAll()" />
      </q-card-actions>
    </div>
  </div>
</template>

<script>
import { ref } from "vue";
import axios from "axios";
import { colors } from "quasar";
import { useQuasar } from "quasar";

const { getPaletteColor } = colors;
const $q = useQuasar();


export default {
  setup() {
    return {
      model: ref(null),
      tab: ref("main"),
      model2: ref(null),
      org: ref(""),
      options: ["xx-xxxx", "xx-xxxx", "xx-xxxx", "xx-xxxx", "xxx-xxx"],
      optionsx: ["111-585758", "xx-xxxx", "xx-xxxx", "xx-xxxx", "9867-28579"],
      code: [
        "A",
        "B",
        "C",
        "D",
        "E",
      ],
      orgs: ["G1", "G2", "G3", "G4"],
      f_type:["1. WOVEN - NYLON, TAFFETA", 
      "2. LIGHT-MIDDLE WEIGHT - SINGLE JERSEY, INTERLOCK, DOUBLE KNIT, JACQUARD",
      "3. FLEECE",
      "4. SPECIAL - RIB, ELASTIC, ELASTAN, LYCRA, SPANDEX, TWO WAY, WAFFLE, PRE WASH, PRESHRUNK, VELOUR"],

      teal: ref(true),
      orange: ref(false),
      red: ref(false),
      cyan: ref(true),
      pink: ref(true),
      basic: ref(false),
      basic2: ref(false),
      text: ref(""),
      SO_YEAR: ref(null),
      SO_NO: ref(null),
      showing: ref(false),
      showing2: ref(false),

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
      status:false,
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
          name: "weightm2",
          label: "Weight M2",
          align: "center",
          field: "weightm2",
        },
        {
          name: "weightyd",
          label: "Weight YD",
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
          name: "lenthyd",
          label: "Lenth YD",
          align: "center",
          field: "lenthyd",
        },
        {
          name: "lenthyds",
          label: "Lenth (YD)",
          align: "center",
          field: "lenthyds",
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
        },
      ],

      columns2: [
        {
          name: "yardroll",
          label: "Issue-Yard/Roll",
          align: "center",
          field: "yardroll",
        },
        {
          name: "lengthyd",
          label: "Length(Yd)",
          align: "center",
          field: "lengthyd",
        },
        {
          name: "marker",
          label: "Obs.Width",
          align: "center",
          field: "marker",
        },
        {
          name: "ttlmarker",
          label: "Ttl Obs.Width",
          align: "center",
          field: "ttlmarker",
        },
        {
          name: "end",
          label: "End1+2",
          align: "center",
          field: "end",
        },
        {
          name: "plug1",
          label: "Plug1",
          align: "center",
          field: "plug1",
        },
        {
          name: "allplug1",
          label: "% of Plug1",
          align: "center",
          field: "allplug1",
        },
        {
          name: "plug2",
          label: "Plug2",
          align: "center",
          field: "plug2",
        },
        {
          name: "allplug2",
          label: "% of Plug2",
          align: "center",
          field: "allplug2",
        },
        {
          name: "plug12",
          label: "Plug12",
          align: "center",
          field: "Plug12",
        },
      ],
      columnsb: [
     {
          name: "endloss",
          label: "Endloss",
          align: "center",
          field: "endloss",
        },
        {
          name: "widthloss",
          label: "Widthloss",
          align: "center",
          field: "widthloss",
        },
        {
          name: "avglenthper",
          label: "AVG_LENGTH%",
          align: "center",
          field: "avglenthper",
        },
        {
          name: "avgend",
          label: "AVG END",
          align: "center",
          field: "vgend",
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
      ],
      comlumnsc: [
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
          label: "Total Cut Out",
          align: "center",
          field: "totalcutout",
        },
        {
          name: "totalcutoutper",
          label: "% of Total Cut Out",
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
          label: "% Remnants Loss",
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
       
      ],
      columns3: [
        {
          name: "no",
          label: "No.",
          align: "center",
          field: "no",
        },
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
          name: "delete",
          label: "Delete",
          align: "center",
          field: "delete",
        },
      ],
      rowwidth: [
        {
          no: "1",
          issue: "",
          avg_width: "",
          delete: "",
        },
        {
          no: "2",
          issue: "",
          avg_width: "",
          delete: "",
        },
        {
          no: "3",
          issue: "",
          avg_width: "",
          delete: "",
        },
        {
          no: "4",
          issue: "",
          avg_width: "",
          delete: "",
        },
        {
          no: "5",
          issue: "",
          avg_width: "",
          delete: "",
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
      line_no: [],
      SO_NO_DOC: "",

      comlumnsmain: [
        {
          name: "issue",
          label: "Issue",
          align: "center",
          field: "issue",
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
          name: "lyd",
          label: "L.YD",
          align: "center",
          field: "lyd",
        },
        {
          name: "linch",
          label: "L.Inch",
          align: "center",
          field: "linch",
        },
        {
          name: "obs",
          label: "Obs.Width",
          align: "center",
          field: "obs",
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
          name: "endplug1",
          label: "End Plug1",
          align: "center",
          field: "endplug1",
        },
        {
          name: "endplug2",
          label: "End Plug2",
          align: "center",
          field: "endplug2",
        },
      ],

      rowmain: [
        {
          issue: "",
          gmm2: "",
          gmyd: "",
          lyd: "",
          linch: "",
          obs: "",
          marker: "",
          ply: "",
          endplug1: "",
          endplug2: "",
        },
      ],
      columnssec: [
        {
          name: "cause",
          label: "Cause",
          align: "center",
          field: "cause",
        },
        {
          name: "code",
          label: "Code",
          align: "center",
          field: "code",
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
          name: "cutoutnoof",
          label: "Cut No of",
          align: "center",
          field: "scutoutnoof",
        },
        {
          name: "cutoutweight",
          label: "Cut No of",
          align: "center",
          field: "scutoutnoof",
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
          name: "avgglueside",
          label: "Avg.Glue Side",
          align: "center",
          field: "avgglueside",
        },
      ],
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
      category$clearable: false,
      category$clear: [],
      mu_seq: "",

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
    
      

      if (this.org == "" || this.F_TYPE == "") {
        this.$q.notify({
          message: "ORG or F TYPE No have Data",
          color: "green-9",
        });
        
      }else{
        if (this.rowx.gm !== "" && this.org !=="") {
        this.isDisabled = true;
        }
      const params = new FormData();
      params.append("GM_NO", this.rowx[index].gm_no);
      params.append("SO_YEAR",this.SO_YEAR);
      params.append("SO_NO",this.SO_NO)
      params.append("CPART_NO",this.cpartno)
      axios({
        method: "post",
        url: this.$api_url + '/somain.php/checkduplicate',
        data: params,
      }).then((resp) => {
        
        if (resp.data.data.length > 0) {
          this.$q.notify({
            position: 'center',
            timeout:1000,
            message: "Duplicate Data Please Try Again",
            color: "red-9",
          });
          
          this.rowx[index].gm_no=""
      
        }else{
          axios({
        method:"post",
        url:this.$api_url + '/somain.php/check_detail',
        data:params
      }).then((resp)=>{
        console.log(resp.data)
        if (resp.data.data.length > 0) {
          
            resp.data.data.forEach((e) => {
               for (let xx = 0; xx <= this.rowx.length; xx++) {
                if (xx == index) {
                    this.rowx[index].gm_no = e.GM_NO;
                    this.rowx[index].GM_M2 = e.GM_M2;
                    this.rowx[index].GM_YD = e.GM_YD;
                    this.rowx[index].WIDTH_INCH = e.WIDTH_INCH;
                    this.rowx[index].YD = e.YD;
                    this.rowx[index].LENGTH_YD = e.LENGTH_YD;
                    this.rowx[index].LENGTH_INCH = e.LENGTH_INCH;
                    this.rowx[index].ttlmarker = (e.YD * e.LAYER).toFixed(2);
                    this.rowx[index].PER_U = e.PER_U;
                    this.rowx[index].LAYER = e.LAYER;
                }
            }
            });
          } 
      }).catch((error)=>{
        console.log(error)
      }) 
        }
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
      this.rowwidth.push([
        {
          no: "",
          issue: "",
          avg_width: "",
          delete: "",
        },
      ]);
    },
    addRowgm() {
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
      console.log(this.rowx);
    },
    removeRow() {
      this.rowwidth.pop();
    },
    Removegm() {
      this.rowx.pop();
    },
    deleteSelected(index) {
      console.log(index);

      this.rowwidth.splice(index, 1);
    },
    deleteRow(index) {
      console.log(index);

      this.rowx.splice(index, 1);
    },
    save() {
      console.log("success");
    },
    search() {
     
  let username = this.$q.localStorage.getItem('username')
      console.log(username)
      console.log(this.SO_YEAR);
      if (this.SO_YEAR == null || this.SO_YEAR == "") {
        this.$q.notify({
          message: "Please Input SO_YEAR",
          color: "green-9",
        });
      }
      console.log(this.SO_NO);
      if (this.SO_NO == null || this.SO_NO == "") {
        this.$q.notify({
          message: "Please Input SO_NO",
          color: "green-9",
        });
      } 
      if (
        this.SO_NO !== "" &&
        this.SO_NO !== null &&
        this.SO_YEAR !== "" &&
        this.SO_YEAR !== ""
      ) {
      }
      const params = new FormData();
      params.append("SO_YEAR", this.SO_YEAR);
      params.append("SO_NO", this.SO_NO);

      axios({
        method: "post",
        url: this.$api_url + '/somain.php/checkso',
        data: params,
      })
        .then((resp) => {
          this.posts = resp.data.data[0];
        })
        .catch((error) => {
          console.log(error);
        });

      axios({
        method: "post",
        url: this.$api_url + '/somain.php/checkcpart',
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

        
    },
    checkgm() {
      const params = new FormData();
      params.append("SO_YEAR", this.SO_YEAR);
      params.append("SO_NO", this.SO_NO);
      params.append("cpartno", this.cpartno);
      this.rowxy=[];
      axios({
        method: "post",
        url: this.$api_url + '/somain.php/checkgm',
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
        }).finally(()=>{
           
        })
    },
    addDetail(index) {
    
    
       let username = this.$q.localStorage.getItem('username')
    
      if (this.exrow == "" || this.rowx[index].number == "") {
        this.$q.notify({
          position:"center",
          timeout:1000,
          size:'xl',
          message: "Please Insert Data",
          color: "green-9",
        });
      } else {
        this.rowwidth = [
          {
            no: "1",
            issue: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "2",
            issue: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "3",
            issue: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "4",
            issue: "",
            avg_width: "",
            delete: "",
          },
          {
            no: "5",
            issue: "",
            avg_width: "",
            delete: "",
          },
        ];
        this.gm_no = this.rowx[index].gm_no
        this.table = this.rowx[index].number;
        this.weightm2 = this.rowx[index].GM_M2;
        this.weightyd = this.rowx[index].GM_YD;
        this.widthinch = this.rowx[index].WIDTH_INCH;
        this.lenthyd = this.rowx[index].YD;
        this.lenthyds = this.rowx[index].LENGTH_YD;
        this.lenthinch = this.rowx[index].LENGTH_INCH;
        this.ttlmarker = this.rowx[index].ttlmarker;
        this.marker = this.rowx[index].PER_U;
        this.ply = this.rowx[index].LAYER;

        this.basic = true;
        console.log(this.gm_no)
         if (this.re == "") {
          axios
            .get(this.$api_url + '/somain.php/createmu')
            .then((resp) => {
            
              this.re.push({
                mu_seq : resp.data.data[0].NEXTVAL
              })

             

                if(this.F_TYPE == "1. WOVEN - NYLON, TAFFETA"){
                  this.F_TYPE = 1
                }
                if(this.F_TYPE == "2. LIGHT-MIDDLE WEIGHT - SINGLE JERSEY, INTERLOCK, DOUBLE KNIT, JACQUARD"){
                  this.F_TYPE = 2
                }
                 if(this.F_TYPE == "3. FLEECE"){
                  this.F_TYPE = 3
                }   
                 if(this.F_TYPE == "4. SPECIAL - RIB, ELASTIC, ELASTAN, LYCRA, SPANDEX, TWO WAY, WAFFLE, PRE WASH, PRESHRUNK, VELOUR"){
                  this.F_TYPE = 4
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
                params.append("SO_YEAR",this.SO_YEAR);
                params.append("SO_NO",this.SO_NO);
                params.append("SO_NO_DOC",this.posts.SO_NO_DOC);
                params.append("STYLE",this.posts.STYLE_REF);
                params.append("CUSTOMER",this.posts.CUST_NAME);
                params.append("QTY",this.posts.QTY);
                params.append('GM_TYPE',this.posts.GMT_TYPE)
                params.append('F_TYPE',this.F_TYPE)
                params.append('FABRIC_TYPE',this.posts.FABRIC_TYPE)
                params.append('DATE',this.date)
                params.append('CPART_NO',this.cpartno)
                params.append('USERNAME' , username)
                for (var pair of params.entries()) {
                  console.log(pair[0] + ", " + pair[1]);
                }

                
                  axios({
                    method:"post",
                    url:this.$api_url + '/somain.php/create_for_head',
                    data:params,
                  }).then((resp)=>{
                    console.log(resp.data)
                     if(resp.data.status == true){
                       axios({
                  method: "post",
                  url: this.$api_url + '/somain.php/createdetail',
                  data: params,
                })
                  .then((resp) => {
                
                    this.stat = resp.data.status;
                   
                    if(this.stat == true){
                  this.rowx[index].status = true;
                    }
                   
                  })
                  .catch((error) => {
                    console.log(error);
                  });
                    }else{
                     
                    } 
                  }).catch((error)=>{
                    console.log(error)
                  }) 
              
               

              } 
            });
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
                params.append("SO_YEAR",this.SO_YEAR);
                params.append("SO_NO",this.SO_NO);
                params.append("SO_NO_DOC",this.posts.SO_NO_DOC);
                params.append("STYLE",this.posts.STYLE_REF);
                params.append("CUSTOMER",this.posts.CUST_NAME);
                params.append("QTY",this.posts.QTY);
                params.append('GM_TYPE',this.posts.GMT_TYPE)
                params.append('FABRIC_TYPE',this.posts.FABRIC_TYPE)
                params.append('DATE',this.date)
                params.append('CPART_NO',this.cpartno)
                params.append('USERNAME' , username)
                for (var pair of params.entries()) {
                  console.log(pair[0] + ", " + pair[1]);
                }
                axios({
                  method: "post",
                  url: this.$api_url + '/somain.php/createdetail',
                  data: params,
                })
                  .then((resp) => {
                
                    this.stat = resp.data.status;
                   
                    if(this.stat == true){
                  this.rowx[index].status = true;
                    }
                  })
                  .catch((error) => {
                    console.log(error);
                  });

     
        } 
      } 
    },
    goDetail(index) {
      this.rowwidth = [];
      this.basic = true;

      this.exrow = this.rowx[index].gm_no;
      console.log(this.exrow);

      const params = new FormData();

      params.append("GM_NO", this.exrow);
      axios({
        method: "post",
        url: this.$api_url + '/somain.php/returndata',
        data: params,
      })
        .then((resp) => {
          console.log(resp.data.data);
          if (resp.data.data.length > 0) {
            resp.data.data.forEach((e) => {
              this.rowwidth.push({
                no: e.LINE_NO,
                issue: e.ISSUE,
                avg_width: e.AVG_WIDTH,
                delete: "",
              }),
                console.log(this.rowwidth);
            });
          }
        })
        .catch((error) => {
          console.log(error);
        });
    },

    Savedata() {
     
      console.log(this.re[0].mu_seq);
       for (var i = 0; i < this.rowwidth.length; i++) {
        const params = new FormData();
        if(this.rowwidth[i].issue !== "" && this.rowwidth[i].avg_width !==""){
        params.append("MU_SEQ", this.re[0].mu_seq);
        params.append("GM_NO", this.exrow);
        params.append("ISSUE", this.rowwidth[i].issue);
        params.append("AVG_WIDTH", this.rowwidth[i].avg_width);
        axios
          .get(this.$api_url + '/somain.php/create_line_no')
          .then((resp) => {
            this.line_no = resp.data.data[0].NEXTVAL;

            params.append("LINE_NO", this.line_no);
            for (var pair of params.entries()) {
              console.log(pair[0] + ", " + pair[1]);
            }

            void axios({
              method: "post",
              url: this.$api_url + '/somain.php/createsubdetail',
              data: params,
            })
              .then((resp) => {
                if (resp.data.status == true) {
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
                  });
                  this.rowwidth = [
                    {
                      no: "1",
                      issue: "",
                      avg_width: "",
                      delete: "",
                    },
                    {
                      no: "2",
                      issue: "",
                      avg_width: "",
                      delete: "",
                    },
                    {
                      no: "3",
                      issue: "",
                      avg_width: "",
                      delete: "",
                    },
                    {
                      no: "4",
                      issue: "",
                      avg_width: "",
                      delete: "",
                    },
                    {
                      no: "5",
                      issue: "",
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
    },
    change_cause(){
      console.log(this.rowmain[0].code)
          if (this.rowmain[0].code == "A") {
          this.rowmain[0].cause = "มาร์คไม่มีผล";
          console.log( this.rowmain[0].cause)
         
        }
        if (this.rowmain[0].code == "B") {
          this.rowmain[0].cause = "ไม่ได้คัดแยกหน้าผ้าปู";
          console.log( this.rowmain[0].cause)
        }
        if (this.rowmain[0].code == "C") {
          this.rowmain[0].cause = "หน้าผ้าไม่สม่ำเสมอในม้วน และต่างม้วน";
          console.log( this.rowmain[0].cause)
        }
        if (this.rowmain[0].code == "D") {
          this.rowmain[0].cause = "ใช้มาร์คไม่เหมาะสมกับหน้าผ้า";
          console.log( this.rowmain[0].cause)
        }
        if (
          this.rowmain[0].code ==
          "E"
        ) {
          this.rowmain[0].cause = "อื่น ๆ ริมเสียบางช่วง, ริมย้วย, ปูผ้าหลายสีหน้าผ้าไม่เท่ากัน";
          console.log( this.rowmain[0].cause)
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

        console.log(this.re[0].mu_seq)
         this.mu_doc_no = this.re[0].mu_seq + this.SO_YEAR + this.SO_NO;
        console.log(this.rowsec[0].code);
        params.append("MU_SEQ", this.re[0].mu_seq);
 
        params.append("SO_YEAR", this.SO_YEAR);
        params.append("SO_NO", this.SO_NO);
        params.append("SO_NO_DOC", this.posts.SO_NO_DOC);
       
        if(this.rowmain[0].endplug1 == undefined){
          this.rowmain[0].endplug1 = ''
        }
        if(this.rowmain[0].endplug2 == undefined){
          this.rowmain[0].endplug2 = ''
        }
        if(this.rowmain[0].code == undefined){
          this.rowmain[0].code = ''
        }
        if(this.rowmain[0].cause == undefined){
          this.rowmain[0].cause = ''
        }
        if(this.rowmain[0].splicenoof == undefined){
          this.rowmain[0].splicenoof = ''
        }
        if(this.rowmain[0].spliceinch == undefined){
          this.rowmain[0].spliceinch = ''
        }
        if(this.rowmain[0].cutoutnoof == undefined){
          this.rowmain[0].cutoutnoof = ''
        }
        if(this.rowmain[0].cutoutweight == undefined){
          this.rowmain[0].cutoutweight = ''
        }
        if(this.rowmain[0].remnants == undefined){
          this.rowmain[0].remnants = ''
        }
        if(this.rowmain[0].widthplug1 == undefined){
          this.rowmain[0].widthplug1 = ''
        }
          if(this.rowmain[0].widthplug2 == undefined){
          this.rowmain[0].widthplug2 = ''
        }
          if(this.rowmain[0].avgglueside == undefined){
          this.rowmain[0].avgglueside = ''
        }

    
        params.append("END_PLUG1", this.rowmain[0].endplug1);
        params.append("END_PLUG2", this.rowmain[0].endplug2);
        params.append("CAUSE_CODE", this.rowmain[0].code);
        params.append("CAUSE_NAME", this.rowmain[0].cause);
        params.append("SPLICE", this.rowmain[0].splicenoof);
        params.append("SPICE_INCH", this.rowmain[0].spliceinch);
        params.append("CUT_OUT", this.rowmain[0].cutoutnoof);
        params.append("CUT_OUT_GRAM", this.rowmain[0].cutoutweight);
        params.append("REMNANTS_WEIGHT", this.rowmain[0].remnants);
        params.append("PLUG1_GRAM", this.rowmain[0].widthplug1);
        params.append("PLUG2_GRAM", this.rowmain[0].widthplug2);
        params.append("GLUE_SIDE", this.rowmain[0].avgglueside);
       
        params.append("MU_DOC_NO", this.mu_doc_no);
        for (var pair of params.entries()) {
          console.log(pair[0] + ", " + pair[1]);
        }

        axios({
          method: "post",
          url: this.$api_url + '/somain.php/update_head',
          data: params,
        })
          .then((resp) => {
            console.log(resp.data)
            
            this.rowmain = [];

            if(resp.data.status == true){
             this.check_bottom_data()
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
    check_bottom_data(){
       const params = new FormData();  
 params.append("MU_SEQ", this.re[0].mu_seq);
 params.append("SO_NO_DOC", this.posts.SO_NO_DOC);
axios({   //all detail ด้านล่าง
          method: "post",
          url: this.$api_url + '/detail.php/all_data_bottom',
          data: params,
        })
          .then((resp) => {
           
            console.log(resp.data)
             if (resp.data.data.length > 0) {
              resp.data.data.forEach((e) => {
                this.rowmain.push({
                    issue: e.ISSUE,
                  gmm2: e.WEIGHT_M2,
                  gmyd: e.WEIGHT_YD,
                  lyd: e.LENGTH_YD,
                  linch: e.LENGTH_INCH,
                  obs: e.OBS_WIDTH,
                  marker: e.MARKER,
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
                  yard_roll:e.YARD_ROLL,
                  length:e.LENGTH_YD,
                  obswidth:e.OBS_WIDTH,
                  ttlobs_widthloss:e.OBS_WIDTH,
                  end12:e.END1_2,
                  plug1:e.PLUG1_INCH,
                  plug2:e.PLUG2_INCH,
                  perplug1:e.PER_PLUG1,
                  perplug2:e.PER_PLUG2,
                  plug12:e.PLUG1_2_INCH,
                  endloss:parseFloat(e.END_LOSS_YD).toFixed(2),
                  widthloss:e.WITH_LOSS,
                  avglenthper:e.AVG_END_LOSS_YD,
                  avgend:e.AVG_END_LOSS,
                  sectionlosslenth:parseFloat(e.SECTION_LOSS_YD).toFixed(2),
                  sectionlossper:e.SECTION_LOSS_PER,
                  spliceloss:e.SPLICE_LOSS_YD,
                  splicelossper:parseFloat(e.SPLICE_LOSS_PER).toFixed(2),
                  overlap:e.OVER_LAB_YD,
                  overlapper:parseFloat(e.OVER_LAB_PER).toFixed(2),
                  defect:parseFloat(e.DEFECT_YD).toFixed(2),
                  defectper:parseFloat(e.DEFECT_PER).toFixed(2),
                  totalcutout: parseFloat(e.TOTAL_CUTOUT_YD).toFixed(2),
                  totalcutoutper:parseFloat(e.TOTAL_CUTOUT_PER).toFixed(2),
                  rement:parseFloat(e.REMENT_LOSS_YD).toFixed(2),
                  rement_per:parseFloat(e.REMENT_LOSS_PER).toFixed(2),
                  totallenthloss:parseFloat(e.TOTAL_LENGTH_LOSS_YD).toFixed(2),
                  totallenthlossper:parseFloat(e.TOTAL_LEN_LOSS_PER).toFixed(2),
                 
                  
                  
                });
              });
            }  
          })
          .catch((error) => {
            console.log(error);
          });
    },
    Back(){
      window.history.back();
    },
  },
  watch: {
    cpartno: function (value) {
      this.checkgm(value);
    },
  },
};
</script>
<style lang="sass">
.my-card
  padding-top: 50px
  padding-right: 30px
  padding-bottom: 50px
  padding-left: 80px
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
