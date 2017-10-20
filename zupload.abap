
report zmmbi0130 .

*----------------------------------------------------------------------*
*- Tipos SAP
*----------------------------------------------------------------------*
type-pools:
  abap .
*----------------------------------------------------------------------*
*- Tabelas
*----------------------------------------------------------------------*

*----------------------------------------------------------------------*
*- Tipos locais
*----------------------------------------------------------------------*

*----------------------------------------------------------------------*
*- Classe
*----------------------------------------------------------------------*

class lcl_class definition .

  public section .

    methods get_data
     importing
       !file        type localfile
       !bukrs       type bukrs
       !lgort       type lgort_d .

    methods show_data .

    class-methods file_open_dialog
      changing
        !file type localfile .

  protected section .

    methods hotspot
      for event if_salv_events_actions_table~link_click
             of cl_salv_events_table
      importing row
                column.

    methods user_command
      for event if_salv_events_functions~added_function
             of cl_salv_events_table
      importing e_salv_function.

    methods progress
      importing
        !percent type i
        !text    type char50 .

  private section .

    types:
      begin of ty_data_tab,
        indice   type char02,
        nftype   type j_1bnftype,
        branch   type j_1bbranc_,
        parid    type j_1bparid,
        docdat   type char08,
        matnr    type matnr,
        werks    type werks_d,
        menge    type char13,
        charg    type charg_d,
        docref   type char10,
        itmref   type char05,
        netpr    type char16,
        cfop     type char6,
        taxlw1   type char03,
        taxlw2   type char03,
        taxtyp   type char04,
        base     type char16,
        rate     type char6,
        excbas   type char16,
        bspiscof type char16,
      end of ty_data_tab,

      begin of ty_mchb,
        matnr type mchb-matnr,
        werks type mchb-werks,
        lgort type mchb-lgort,
        charg type mchb-charg,
        clabs type mchb-clabs,
      end of ty_mchb,

      begin of ty_out,
        mblnr    type rm07m-mblnr,
        mjahr    type rm07m-mjahr,
        zeile    type mseg-zeile,
        docnum   type j_1bnfdoc-docnum,
        werks    type mchb-werks,
        matnr    type mchb-matnr,
        lgort    type mchb-lgort,
        charg    type mchb-charg,
        clabs    type mchb-clabs,
        status   type char1,
        descri   type char100,
        indice   type char02,
        itmnum   type itmnum,
        nftype   type j_1bnfdoc-nftype,
        branch   type j_1bnfdoc-branch,
        parid    type j_1bnfdoc-parid,
        docdat   type j_1bnfdoc-docdat,
        matn1    type j_1bnflin-matnr,
        werk1    type j_1bnflin-werks,
        menge    type j_1bnflin-menge,
        char1    type j_1bnflin-charg,
        docref   type j_1bnflin-docref,
        itmref   type j_1bnflin-itmref,
        netpr    type j_1bnflin-netpr,
        cfop     type j_1bnflin-cfop,
        taxlw1   type j_1bnflin-taxlw1,
        taxlw2   type j_1bnflin-taxlw2,
        taxtyp   type j_1bnfstx-taxtyp,
        base     type j_1bnfstx-base,
        rate     type j_1bnfstx-base,
        excbas   type j_1bnfstx-base,
        bspiscof type j_1bnfstx-base,
        color    type lvc_t_scol,
      end of ty_out,

      begin of ty_mara,
        matnr type mara-matnr,
        matkl type mara-matkl,
        meins type mara-meins,
        plgtp type mara-plgtp,
      end of ty_mara,

      begin of ty_makt,
        matnr type makt-matnr,
        maktx type makt-maktx,
      end of ty_makt,

      begin of ty_marc,
        matnr type marc-matnr,
        werks type marc-werks,
        steuc type marc-steuc,
      end of ty_marc,

      begin of ty_mbew,
        matnr type mbew-matnr,
        bwkey type mbew-bwkey,
        bwtar type mbew-bwtar,
        mtuse type mbew-mtuse,
        mtorg type mbew-mtorg,
      end of ty_mbew,

      begin of ty_mseg,
        mblnr type mseg-mblnr,
        mjahr type mseg-mjahr,
        zeile type mseg-zeile,
        matnr type mseg-matnr,
      end of ty_mseg,

      tab_data_tab type table of ty_data_tab,
      tab_mchb     type table of ty_mchb,
      tab_out      type table of ty_out,
      tab_mara     type table of ty_mara,
      tab_makt     type table of ty_makt,
      tab_marc     type table of ty_marc,
      tab_mbew     type table of ty_mbew,
      tab_mseg     type table of ty_mseg.


    data:
      out    type tab_out,
      mchb   type tab_mchb,
      j_1baa type table of j_1baa,
      j_1bb2 type table of j_1bb2,
      bukrs  type bukrs,
      mara   type tab_mara,
      makt   type tab_makt,
      marc   type tab_marc,
      mbew   type tab_mbew,
      table  type ref to cl_salv_table,
      return type bapiret2_t .

    methods get_file
      importing
        !file type localfile
      exporting
        !data_tab type tab_data_tab .

    methods search
      importing
        !data_tab type tab_data_tab
        !bukrs    type bukrs
        !lgort    type lgort_d
      exporting
        !mchb     type tab_mchb .

    methods organize
      importing
        !data_tab type tab_data_tab
        !mchb     type tab_mchb .

    methods criar_baixa_estoque
      importing
        data             type tab_out
      exporting
        return           type bapiret2_t .

    methods criar_nf_manual
      importing
        data             type tab_out
      exporting
        return           type bapiret2_t .

    methods config_status
      changing
        !table type ref to cl_salv_table .

    methods config_layout
      changing
        !table type ref to cl_salv_table .

    methods config_column
      changing
        !table type ref to cl_salv_table .

    methods config_description
      importing
        !field type lvc_fname
        !text  type scrtext_l
      changing
        !table type ref to cl_salv_table .

    methods process .
    methods process_pai .

    methods convert_num
      importing
        !char type any
      returning
        value(numc) type j_1bnetpri .

    methods convert_dats
      importing
        !char type any
      returning
        value(date) type sydatum .

    methods convert_matnr
      changing
        !matnr type any .

    methods convert_branch
      changing
        !branch type any .

    methods convert_parid
      changing
        !parid type any .

    methods update_docnum
      importing
        !indice type ty_data_tab-indice
        !docnum type j_1bdocnum .

    methods update_mblnr
      importing
        !indice type ty_data_tab-indice
        !mblnr  type mblnr
        !mjahr  type mjahr .

    methods question
      returning
        value(answer) type char1 .

    methods get_j_1bnfdoc
      importing
        !data      type ty_out
      changing
        !j_1bnfdoc type bapi_j_1bnfdoc
        !return    type bapiret2_t .

    methods get_j_1bnflin
      importing
        !data      type ty_out
      changing
        !j_1bnflin type bapi_j_1bnflin_tab
        !return    type bapiret2_t .


    methods get_j_1bnfstx
      importing
        !data      type ty_out
      changing
        !j_1bnfstx type bapi_j_1bnfstx_tab
        !return    type bapiret2_t .

    methods bapiret
      importing
        !type       type bapiret2-type
        !message_v1 type bapiret2-message_v1
        !message_v2 type bapiret2-message_v2 optional
        !message_v3 type bapiret2-message_v3 optional
        !message_v4 type bapiret2-message_v4 optional
      changing
        !return     type bapiret2_t .

    methods get_zeile
      changing
        !data type tab_out .

    methods get_percent
      importing
        !total  type i
        !indice type any
      returning
        value(percent) type i .

endclass .                    "lcl_class DEFINITION

*----------------------------------------------------------------------*
*       CLASS lcl_class IMPLEMENTATION
*----------------------------------------------------------------------*
class lcl_class implementation .

  method get_data .

    data:
      data_tab type me->tab_data_tab,
      lt_mchb  type me->tab_mchb .

    me->progress(
      exporting
        percent = 10
        text    = 'Importando arquivo.'
    ).

    me->get_file(
      exporting
        file     = file
      importing
        data_tab = data_tab
    ).

    me->progress(
      exporting
        percent = 30
        text    = 'Buscando informações.'
    ).

    me->search(
      exporting
        data_tab = data_tab
        bukrs    = bukrs
        lgort    = lgort
      importing
        mchb     = lt_mchb
    ) .

    me->progress(
      exporting
        percent = 60
        text    = 'Organizando dados.'
    ).

    me->organize(
      exporting
        data_tab = data_tab
        mchb     = lt_mchb
    ).

  endmethod .                    "get_data

  method show_data .

    data:
      display    type ref to cl_salv_display_settings,
      events     type ref to cl_salv_events_table,
      selections type ref to cl_salv_selections .

    if lines( out ) eq 0 .

    else .

      try.

          call method cl_salv_table=>factory
            importing
              r_salv_table = table
            changing
              t_table      = out.

*         Eventos do relatório
          events = table->get_event( ).
          set handler me->user_command for events.
          set handler me->hotspot      for events.

          data lo_sel type ref to cl_salv_selections.
          lo_sel = table->get_selections( ).
          lo_sel->set_selection_mode( if_salv_c_selection_mode=>multiple ).


          me->config_status(
            changing
              table = table
          ).

*         Configurando Layout
          me->config_layout(
            changing
              table = table
          ) .

*         Configurando Colunas
          me->config_column(
            changing
              table = table
          ) .

*         Layout de Zebra
          display = table->get_display_settings( ) .
          display->set_striped_pattern( cl_salv_display_settings=>true ) .

*         Enable cell selection mode
          selections = table->get_selections( ).
          selections->set_selection_mode( if_salv_c_selection_mode=>row_column ).

          table->display( ).

        catch cx_salv_msg .
        catch cx_salv_not_found .
        catch cx_salv_existing .
        catch cx_salv_data_error .
        catch cx_salv_object_not_found .

      endtry.

    endif .


  endmethod .                    "show_data

  method file_open_dialog .

    data:
      file_table type filetable,
      line       type file_table,
      rc         type i .

    clear:
      file .

    call method cl_gui_frontend_services=>file_open_dialog
      exporting
*       window_title            =
*       default_extension       =
*       default_filename        =
        file_filter             = cl_gui_frontend_services=>filetype_excel
*       with_encoding           =
*       initial_directory       =
*       multiselection          =
      changing
        file_table              = file_table
        rc                      = rc
*       user_action             =
*       file_encoding           =
      exceptions
        file_open_dialog_failed = 1
        cntl_error              = 2
        error_no_gui            = 3
        not_supported_by_gui    = 4
        others                  = 5 .

    if sy-subrc eq 0 .

      read table file_table into line index 1 .
      if sy-subrc eq 0 .
        file = line-filename .
      endif .

    else .

      message id sy-msgid type sy-msgty number sy-msgno
            with sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.

    endif.


  endmethod .                    "file_open_dialog

  method hotspot .

    data:
      line type me->ty_out .

    read table me->out into line index row .

    if sy-subrc eq 0 .

      case column .

        when 'DOCNUM' .

          set parameter id 'JEF' field line-docnum .
          call transaction 'J1B3N' and skip first screen.

        when 'MBLNR' .

          set parameter id 'MBN' field line-mblnr .
          set parameter id 'MJA' field line-mjahr .
          call transaction 'MB03' and skip first screen.

        when others .

      endcase .

    endif.


  endmethod .                    "hotspot

  method user_command .

    case e_salv_function .

      when 'BAIXA' .
        me->process( ) .

      when others .

    endcase .


    if table is bound .

      me->process_pai( ) .

      table->refresh( ) .

    endif .

  endmethod .                    "user_command

  method progress .

    data:
      message type char50 .

    if ( percent is not initial ) and
       ( text    is not initial ) .

      clear:
        message .

      concatenate text
                  'Aguarde...'
             into message
        separated by space .

      call function 'SAPGUI_PROGRESS_INDICATOR'
        exporting
          percentage = percent
          text       = message.

    endif .

  endmethod .                    "progress

  method get_file .

    data:
      filename  type rlgrap-filename,
      intern    type table of kcde_cells,
      line      type kcde_cells,
      row       type numc4,
      data_line type me->ty_data_tab .

    field-symbols:
      <field> type any .

    refresh:
      data_tab .

    filename = file .

    call function 'KCD_EXCEL_OLE_TO_INT_CONVERT'
      exporting
        filename                = filename
        i_begin_col             = 1
        i_begin_row             = 2
        i_end_col               = 20
        i_end_row               = 9999
      tables
        intern                  = intern
      exceptions
        inconsistent_parameters = 1
        upload_ole              = 2
        others                  = 3.

    if sy-subrc eq 0 .

      loop at intern into line .

        if row is initial .
          row = line-row .
        else .
          if line-row ne row .
            row = line-row .

*           Convertendo Valores
            me->convert_matnr(
              changing
                matnr = data_line-matnr
            ) .

            me->convert_branch(
              changing
                branch = data_line-branch
            ) .

            me->convert_parid(
              changing
                parid = data_line-parid
            ) .

            append data_line to data_tab .
            clear  data_line .
          endif .
        endif .

        assign component line-col of structure data_line to <field> .
        if <field> is assigned.
          <field> = line-value .
          unassign <field> .
        endif.

      endloop .

    else .
      message id sy-msgid type sy-msgty number sy-msgno
            with sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    endif.

    free:
      intern .

  endmethod .                    "get_file

  method search .

    data:
      lt_data type me->tab_data_tab .

    field-symbols:
      <data> type me->ty_data_tab .

    refresh:
      mchb .

    append lines of data_tab to lt_data .
    sort lt_data by:
      matnr ascending
      werks ascending
      charg ascending .
    delete adjacent duplicates from lt_data
    comparing matnr werks charg .

    if lines( lt_data ) eq 0 .
    else .
    endif .

    if lgort is initial .

      select matnr werks lgort charg clabs
        into table mchb
        from mchb
         for all entries in lt_data
       where matnr eq lt_data-matnr
         and werks eq lt_data-werks
         and charg eq lt_data-charg .

    else .

      select matnr werks lgort charg clabs
        into table mchb
        from mchb
         for all entries in lt_data
       where matnr eq lt_data-matnr
         and werks eq lt_data-werks
         and lgort eq lgort
         and charg eq lt_data-charg .

    endif .

*   Dados para crição de NF
    refresh:
      lt_data, me->j_1baa, me->j_1bb2 .
    append lines of data_tab to lt_data .
    sort lt_data by nftype ascending .
    delete adjacent duplicates from lt_data
    comparing nftype .

    select *
      into table me->j_1baa
      from j_1baa
       for all entries in lt_data
     where nftype eq lt_data-nftype.

    refresh: me->j_1bb2 .
    append lines of data_tab to lt_data .
    sort lt_data by branch ascending .
    delete adjacent duplicates from lt_data
    comparing branch .

    select *
      into table me->j_1bb2
      from j_1bb2
       for all entries in lt_data
     where bukrs  eq bukrs
       and branch eq lt_data-branch .

    me->bukrs = bukrs .


*   Dados de Item de NF

    refresh:
      me->mara, me->makt, lt_data .
    append lines of data_tab to lt_data .
    sort lt_data by matnr ascending .
    delete adjacent duplicates from lt_data
    comparing matnr .

    if lines( lt_data ) ne 0 .

      select matnr matkl meins plgtp
        into table me->mara
        from mara
         for all entries in lt_data
       where matnr eq lt_data-matnr .

      select matnr maktx
        into table me->makt
        from makt
         for all entries in lt_data
       where spras eq sy-langu
         and matnr eq lt_data-matnr .

    endif .

    refresh:
      me->marc, me->marc, lt_data .
    append lines of data_tab to lt_data .
    sort lt_data by matnr werks ascending .
    delete adjacent duplicates from lt_data
    comparing matnr werks .

    select matnr werks steuc
      into table me->marc
      from marc
       for all entries in lt_data
     where matnr eq lt_data-matnr
       and werks eq lt_data-werks .

    select matnr bwkey bwtar mtuse mtorg
      into table me->mbew
      from mbew
       for all entries in lt_data
     where matnr eq lt_data-matnr
       and bwkey eq lt_data-werks .

    free:
      lt_data .

  endmethod .                    "search

  method organize .

    data:
      data_line  type ty_data_tab,
      mchb_line  type ty_mchb,
      out_line   type ty_out,
      color_line type lvc_s_scol,
      saldo      type mchb-clabs .

    loop at data_tab into data_line .

      read table mchb into mchb_line
        with key matnr = data_line-matnr
                 werks = data_line-werks
                 charg = data_line-charg .
      if sy-subrc eq 0 .

*       Configuração de Status
        saldo = mchb_line-clabs - data_line-menge .
        if saldo ge 0 .
          out_line-status = '1' .
        else .

          out_line-status = '0' .
          write saldo to out_line-descri decimals 2 .
          condense out_line-descri no-gaps .
          concatenate 'Saldo insuficiente:'
                      out_line-descri
                 into out_line-descri
            separated by space .

          color_line-color-col = 6.
          color_line-color-int = 0.
          color_line-color-inv = 0.
          append color_line to out_line-color .
          clear  color_line .
        endif .

        out_line-werks = mchb_line-werks .
        out_line-matnr = mchb_line-matnr .
        out_line-lgort = mchb_line-lgort .
        out_line-charg = mchb_line-charg .
        out_line-clabs = mchb_line-clabs .

      endif .
      out_line-indice   = data_line-indice .
      out_line-nftype   = data_line-nftype .
      out_line-branch   = data_line-branch .
      out_line-parid    = data_line-parid .
      out_line-docdat   = me->convert_dats( data_line-docdat ) .
      out_line-matn1    = data_line-matnr .
      out_line-werk1    = data_line-werks .
      out_line-menge    = data_line-menge .
      out_line-char1    = data_line-charg .
      out_line-docref   = data_line-docref .
      out_line-itmref   = data_line-itmref .

      out_line-netpr    = me->convert_num( data_line-netpr ) .
      out_line-cfop     = data_line-cfop .
      out_line-taxlw1   = data_line-taxlw1 .
      out_line-taxlw2   = data_line-taxlw2 .
      out_line-taxtyp   = data_line-taxtyp .
      out_line-base     = data_line-base .
      out_line-rate     = data_line-rate .
      out_line-excbas   = me->convert_num( data_line-excbas ) .
      out_line-bspiscof = me->convert_num( data_line-bspiscof ) .

      append out_line to out .
      clear  out_line .

    endloop .


  endmethod .                    "organize


  method criar_baixa_estoque .

    data:
      goodsmvt_header  type bapi2017_gm_head_01,
      goodsmvt_code    type bapi2017_gm_code,
      goodsmvt_item    type tab_bapi_goodsmvt_item,
      materialdocument type bapi2017_gm_head_ret-mat_doc,
      matdocumentyear  type bapi2017_gm_head_ret-doc_year,
      data_line        type me->ty_out,
      item             type bapi2017_gm_item_create .

    refresh:
      return .

    goodsmvt_header-pstng_date = sy-datum .
    goodsmvt_header-doc_date   = sy-datum .
    goodsmvt_header-header_txt = 'Baixa de estoque/emissão NF' .
    goodsmvt_code              = '05' .

    loop at data into data_line .

      item-material   = data_line-matnr .
      item-plant      = data_line-werks .
      item-stge_loc   = data_line-lgort .
      item-batch      = data_line-charg .
      item-move_type  = 'Z22' .
*     item-entry_qnt  = data_line-clabs .
      item-entry_qnt  = data_line-menge .

      append item to goodsmvt_item .
      clear  item .

    endloop .

    if lines( goodsmvt_item ) eq 0 .

    else .

      call function 'BAPI_GOODSMVT_CREATE'
        exporting
          goodsmvt_header               = goodsmvt_header
          goodsmvt_code                 = goodsmvt_code
          testrun                       = if_salv_c_bool_sap=>false
*         testrun                       = if_salv_c_bool_sap=>true
*         goodsmvt_ref_ewm              =
        importing
*         goodsmvt_headret              =
          materialdocument              = materialdocument
          matdocumentyear               = matdocumentyear
        tables
          goodsmvt_item                 = goodsmvt_item
*         goodsmvt_serialnumber         =
          return                        = return
*         goodsmvt_serv_part_data       =
*         extensionin                   =
                .

      read table return transporting no fields
        with key type = 'E' .

      if sy-subrc eq 0 .

      else .

        call function 'BAPI_TRANSACTION_COMMIT'
          exporting
            wait          = if_salv_c_bool_sap=>true
*         importing
*           return        =
          .

        me->update_mblnr(
          indice = data_line-indice
          mblnr  = materialdocument
          mjahr  = matdocumentyear
        ) .

      endif .

    endif .

  endmethod .                    "realizar_baixa


  method criar_nf_manual .

    data:
      obj_header   type bapi_j_1bnfdoc,
      nfcheck      type bapi_j_1bnfcheck,
      obj_item     type bapi_j_1bnflin_tab,
      item         type bapi_j_1bnflin,
      obj_item_tax type bapi_j_1bnfstx_tab,
      line         type me->ty_out,
      docnum       type bapi_j_1bnfdoc-docnum,
      itmnum       type itmnum .


    read table data into line index 1 .

    if sy-subrc eq 0 .

      clear:
        obj_header, itmnum .
      refresh:
        obj_item ,obj_item_tax, return .

*     Monta opçes de veridficação na criação da NF
*     12345678901234567890123456789012345
      nfcheck = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' .

      if obj_header is initial .

        me->get_j_1bnfdoc(
          exporting
            data      = line
          changing
            j_1bnfdoc = obj_header
            return    = return
        ).

      endif .

    endif .


    read table return transporting no fields
      with key type = 'E' .

    if sy-subrc eq 0 .

      append lines of return to me->return .

    else .

      loop at data into line .

        me->get_j_1bnflin(
          exporting
            data      = line
          changing
            j_1bnflin = obj_item
            return    = return
        ).


        me->get_j_1bnfstx(
          exporting
            data      = line
          changing
            j_1bnfstx = obj_item_tax
            return    = return
        ).

      endloop.

    endif .


    if lines( obj_item ) gt 0 .

      call function 'BAPI_J_1B_NF_CREATEFROMDATA'
        exporting
          obj_header           = obj_header
*         obj_header_add       =
          nfcheck              = nfcheck
        importing
          e_docnum             = docnum
        tables
*         obj_partner          =
          obj_item             = obj_item
*         obj_item_add         =
          obj_item_tax         = obj_item_tax
*         obj_header_msg       =
*         obj_refer_msg        =
*         obj_ot_partner       =
*         obj_import_di        =
*         obj_import_adi       =
          return               = return .

      read table return transporting no fields
        with key type = 'E' .

      if sy-subrc eq 0 .

      else .

        call function 'BAPI_TRANSACTION_COMMIT'
          exporting
            wait          = if_salv_c_bool_sap=>true
*         importing
*           return        =
                  .

        me->update_docnum(
          indice = line-indice
          docnum = docnum
        ) .

      endif .

    endif .


  endmethod .                    "cria_doc_material

  method config_status .

    data:
      functions type ref to cl_salv_functions,
      func_list type salv_t_ui_func,
      list_line like line of func_list .

    try .

        table->set_screen_status(
          pfstatus      = 'STANDARD_FULLSCREEN'
          report        = sy-cprog
          set_functions = table->c_functions_all
        ).

*       Verificar se algum item não tem saldo suficiente
        read table me->out transporting no fields
          with key status = '0' .
        if sy-subrc eq 0 .
        else .
*         Verifica se os itens ja foram processados
          if lines( me->return ) gt 0 .
            read table me->return transporting no fields
              with key type = 'S' .
            if sy-subrc eq 0 .
            endif .
          endif .
        endif .

        if sy-subrc eq 0 .

          functions =   table->get_functions( ).
          func_list = functions->get_functions( ).

          loop at func_list into list_line.
            if list_line-r_function->get_name( ) = 'BAIXA'.
              list_line-r_function->set_visible( if_salv_c_bool_sap=>false ).
            endif.
          endloop.

        endif .

      catch cx_salv_not_found .
      catch cx_salv_wrong_call .

    endtry .


  endmethod .                    "config_status


  method config_layout .

    data:
      layout type ref to cl_salv_layout,
      key    type salv_s_layout_key .

    layout     = table->get_layout( ).
    key-report = sy-repid.
    layout->set_key( key ).
    layout->set_save_restriction( if_salv_c_layout=>restrict_none ).

  endmethod .                    "config_layout

  method config_column .

    data:
      column  type ref to cl_salv_column_list,
      columns type ref to cl_salv_columns_table .

    try .

        columns = table->get_columns( ).
        columns->set_optimize( if_salv_c_bool_sap=>true ).

        columns->set_color_column( 'COLOR' ).

        column ?= columns->get_column( 'ITMNUM' ).
        column->set_visible( if_salv_c_bool_sap=>false ) .

        column ?= columns->get_column( 'DOCNUM' ).
        column->set_visible( if_salv_c_bool_sap=>false ) .
        column ?= columns->get_column( 'MBLNR' ).
        column->set_visible( if_salv_c_bool_sap=>false ) .
        column ?= columns->get_column( 'MJAHR' ).
        column->set_visible( if_salv_c_bool_sap=>false ) .
        column ?= columns->get_column( 'ZEILE' ).
        column->set_visible( if_salv_c_bool_sap=>false ) .

        me->config_description(
          exporting
            field = 'STATUS'
            text  = 'Status'
          changing
            table = table
        ) .

        me->config_description(
          exporting
            field = 'DESCRI'
            text  = 'Descrição do Status'
          changing
            table = table
        ) .

        me->config_description(
          exporting
            field = 'CLABS'
            text  = 'Quantidade'
          changing
            table = table
        ) .

        me->config_description(
          exporting
            field = 'INDICE'
            text  = 'Quebra NF'
          changing
            table = table
        ) .

        me->config_description(
          exporting
            field = 'PARID'
            text  = 'Fornecedor'
          changing
            table = table
        ) .

        me->config_description(
          exporting
            field = 'DOCDAT'
            text  = 'Data NF'
          changing
            table = table
        ) .

        me->config_description(
          exporting
            field = 'NETPR'
            text  = 'Preço'
          changing
            table = table
        ) .

        me->config_description(
          exporting
            field = 'TAXLW1'
            text  = 'Cód D.F ICMS'
          changing
            table = table
        ) .

        me->config_description(
          exporting
            field = 'TAXLW2'
            text  = 'Cód IPI'
          changing
            table = table
        ) .

        me->config_description(
          exporting
            field = 'TAXTYP'
            text  = 'Cód ICMS'
          changing
            table = table
        ) .

        me->config_description(
          exporting
            field = 'BASE'
            text  = 'Cód D.F ICMS'
          changing
            table = table
        ) .

        me->config_description(
          exporting
            field = 'RATE'
            text  = 'Cód IPI'
          changing
            table = table
        ) .

        me->config_description(
          exporting
            field = 'EXCBAS'
            text  = 'Cód ICMS'
          changing
            table = table
        ) .

        me->config_description(
          exporting
            field = 'BSPISCOF'
            text  = 'Cód ICMS'
          changing
            table = table
        ) .

      catch cx_salv_data_error .

      catch cx_salv_not_found .

    endtry .

  endmethod .                    "config_column

  method config_description .

    data:
      long_text   type scrtext_l,
      medium_text type scrtext_m,
      short_text  type scrtext_s,
      column      type ref to cl_salv_column_list,
      columns     type ref to cl_salv_columns_table .

    try .

        columns = table->get_columns( ).

        short_text = medium_text = long_text = text .

        if columns is bound .

          column ?= columns->get_column( field ).

          if column is bound .

            column->set_long_text( long_text ).
            column->set_medium_text( medium_text ).
            column->set_short_text( short_text ).

          endif .

        endif .

      catch cx_salv_not_found .

    endtry .

    free:
      column, columns .

  endmethod .                    "config_description

  method process .

    data:
      materialdocument type bapi2017_gm_head_ret-mat_doc,
      matdocumentyear  type bapi2017_gm_head_ret-doc_year,
      bapiret2         type bapiret2_t,
      itmnum           type itmnum,
      indice           type me->tab_out,
      indice_line      type me->ty_out,
      data             type me->tab_out,
      out_line         type me->ty_out,
      total            type i,
      percent          type i .

    if me->question( ) eq '1' .

      refresh:
        me->return, bapiret2 .

      append lines of me->out to indice .
      sort indice ascending by indice .
      delete adjacent duplicates from indice comparing indice .

      total = lines( indice ) .

      loop at indice into indice_line .

        percent =
          me->get_percent(
            total  = total
            indice = sy-tabix
          ).

        me->progress(
          percent = percent
          text    = 'Criando Baixa e NF Manual.'
        ) .

        loop at me->out into out_line
                       where indice eq indice_line-indice .
          itmnum          = itmnum + 10 .
          out_line-itmnum = itmnum .
          append out_line to data .
          clear  out_line .
        endloop .

        if lines( data ) gt 0 .

          me->criar_baixa_estoque(
            exporting
              data             = data
            importing
              return           = bapiret2
          ) .

          if lines( bapiret2 ) gt 0 .
            append lines of bapiret2 to me->return .
          endif .

          me->get_zeile(
            changing
              data = data
          ) .

          me->criar_nf_manual(
            exporting
              data             = data
            importing
              return           = bapiret2
          ) .

          if lines( bapiret2 ) gt 0 .
            append lines of bapiret2 to me->return .
          endif .

          clear:
            itmnum .

          refresh:
            data, bapiret2 .

        endif .


      endloop .

      if lines( me->return ) gt 0 .

        call function 'FINB_BAPIRET2_DISPLAY'
          exporting
            it_message = me->return.

      endif .

    else .

    endif.


    me->config_status(
      changing
        table = me->table
   ) .


  endmethod .                    "link_click

  method process_pai .

    data:
      column  type ref to cl_salv_column_list,
      color   type lvc_s_colo,
      columns type ref to cl_salv_columns_table .

    if table is bound .

      try .

          columns = table->get_columns( ).
          columns->set_optimize( if_salv_c_bool_sap=>true ).

          column ?= columns->get_column( 'DOCNUM' ).
*         Alterando Cor de exibição
          color-col = 5.
          color-int = 0.
          color-inv = 0.
          column->set_color( color ) .
*         Exibindo novamente a Coluna
          column->set_visible( if_salv_c_bool_sap=>true ) .
*         Setando hotspot
          column->set_cell_type( if_salv_c_cell_type=>hotspot ).
*         Posicionando
          columns->set_column_position(
            columnname = 'DOCNUM'
            position   = 8
          ) .

          column ?= columns->get_column( 'MBLNR' ).
          color-col = 5.
          color-int = 0.
          color-inv = 0.
          column->set_color( color ) .
          column->set_visible( if_salv_c_bool_sap=>true ) .
          column->set_cell_type( if_salv_c_cell_type=>hotspot ).
          columns->set_column_position(
            columnname = 'MBLNR'
            position   = 8
          ) .

*         Ocultando algumas colunas para que fiquei mais "clean" o relatório
          column ?= columns->get_column( 'STATUS' ).
          column->set_visible( if_salv_c_bool_sap=>false ) .
          column ?= columns->get_column( 'DESCRI' ).
          column->set_visible( if_salv_c_bool_sap=>false ) .
          column ?= columns->get_column( 'CLABS' ).
          column->set_visible( if_salv_c_bool_sap=>false ) .

        catch cx_salv_not_found .

      endtry .

    endif .

  endmethod .                    "process_pai

  method convert_num .

    if char is not initial .

      call function 'CHAR_FLTP_CONVERSION'
        exporting
*         dyfld                    = ' '
*         maskn                    = ' '
*         maxdec                   = '16'
*         maxexp                   = '59+'
*         minexp                   = '60-'
          string                   = char
*         msgtyp_decim             = 'w'
        importing
*         decim                    =
*         expon                    =
          flstr                    = numc
*         ivalu                    =
       exceptions
         exponent_too_big         = 1
         exponent_too_small       = 2
         string_not_fltp          = 3
         too_many_decim           = 4
         others                   = 5 .

      if sy-subrc eq 0 .
      else .
        message id sy-msgid type sy-msgty number sy-msgno
              with sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      endif.

    endif .

  endmethod .                    "convert_num


  method convert_dats .

    if char is not initial .

      call function 'CONVERSION_EXIT_PDATE_INPUT'
        exporting
          input  = char
        importing
          output = date.

    endif .

  endmethod .                    "convert_dats


  method convert_matnr .

    if matnr is not initial .

      call function 'CONVERSION_EXIT_MATN1_INPUT'
        exporting
          input        = matnr
        importing
          output       = matnr
        exceptions
          length_error = 1
          others       = 2.

      if sy-subrc eq 0 .

      else .

        message id sy-msgid type sy-msgty number sy-msgno
              with sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.

      endif.


    endif .

  endmethod .                    "convert_matnr

  method convert_branch .

    if branch is not initial .

      call function 'CONVERSION_EXIT_ALPHA_INPUT'
        exporting
          input  = branch
        importing
          output = branch.

    endif .

  endmethod .                    "convert_branch

  method convert_parid .

    if parid is not initial .

      call function 'CONVERSION_EXIT_ALPHA_INPUT'
        exporting
          input  = parid
        importing
          output = parid.

    endif .

  endmethod .                    "convert_parid

  method update_docnum .

    field-symbols:
      <out> type me->ty_out .

    if docnum is not initial .

      loop at me->out assigning <out>
                          where indice eq indice .
        <out>-docnum = docnum .
      endloop .

      unassign:
        <out> .

    endif .

  endmethod .                    "update_docnum

  method update_mblnr .

    field-symbols:
      <out> type me->ty_out .

    if ( mblnr is not initial ) and
       ( mjahr is not initial ) .

      loop at me->out assigning <out>
                          where indice eq indice .
        <out>-mblnr = mblnr .
        <out>-mjahr = mjahr .
      endloop .

      unassign:
        <out> .
    endif .

  endmethod .                    "update_mblnr

  method question .

    data:
      text_question type char100 .


    concatenate 'A ação não poderá ser desfeita.'
                'Tem certeza que deseja'
                'executar a baixa?'
           into text_question
      separated by space .

    call function 'POPUP_TO_CONFIRM'
      exporting
        titlebar                    = 'Confirmação de Baixa'
*       diagnose_object             = ' '
        text_question               = text_question
*       text_button_1               = 'ja'(001)
*       icon_button_1               = ' '
*       text_button_2               = 'nein'(002)
*       icon_button_2               = ' '
*       default_button              = '1'
*       display_cancel_button       = 'x'
*       userdefined_f1_help         = ' '
*       start_column                = 25
*       start_row                   = 6
*       popup_type                  =
*       iv_quickinfo_button_1       = ' '
*       iv_quickinfo_button_2       = ' '
      importing
        answer                      = answer
*     tables
*       parameter                   =
     exceptions
       text_not_found              = 1
       others                      = 2 .

    if sy-subrc eq 0 .

    else .

    endif.


  endmethod .                    "question


  method get_j_1bnfdoc .

    data:
      j_1baa_line type j_1baa,
      j_1bb2_line type j_1bb2 .

    clear:
      j_1bnfdoc .
    refresh:
      return .

    j_1bnfdoc-nftype = data-nftype .

    read table me->j_1baa into j_1baa_line
      with key nftype = data-nftype .

    if sy-subrc eq 0 .

      j_1bnfdoc-doctyp = j_1baa_line-doctyp .
      j_1bnfdoc-direct = j_1baa_line-direct .
      j_1bnfdoc-form   = j_1baa_line-form .
      j_1bnfdoc-model  = j_1baa_line-model .

      read table me->j_1bb2 into j_1bb2_line
        with key bukrs  = me->bukrs
                 branch = data-branch
                 form   = j_1baa_line-form .

      if sy-subrc eq 0 .

        j_1bnfdoc-series = j_1bb2_line-series .

      endif .

    else .

      me->bapiret(
        exporting
          type       = 'E'
          message_v1 = 'Falha ao buscar informação de'
          message_v2 = 'Categoria de Nota Fiscal(J_1BAA)'
        changing
          return     = return
      ).

    endif .

    j_1bnfdoc-docdat = data-docdat .
    j_1bnfdoc-pstdat = sy-datum .
    j_1bnfdoc-credat = sy-datum .
    j_1bnfdoc-cretim = sy-uzeit .
    j_1bnfdoc-crenam = sy-uname .
    j_1bnfdoc-manual = if_salv_c_bool_sap=>true .
    j_1bnfdoc-waerk  = 'BRL'.
    j_1bnfdoc-bukrs  = me->bukrs .
    j_1bnfdoc-branch = data-branch .
    j_1bnfdoc-nfe    = j_1baa_line-nfe .
    j_1bnfdoc-parvw  = 'LF' .
    j_1bnfdoc-parid = data-parid .

  endmethod .                    "get_j_1bnfdoc


  method get_j_1bnflin .

    data:
      line       type bapi_j_1bnflin,
      mara_line  type me->ty_mara,
      makt_line  type me->ty_makt,
      marc_line  type me->ty_marc,
      mbew_line  type me->ty_mbew,
      message    type symsgv.


    line-itmnum = data-itmnum .
    line-matnr  = data-matnr .
    line-reftyp = 'MD' .
    concatenate data-mblnr data-mjahr into line-refkey .
    line-refitm = data-zeile .


    read table me->mara into mara_line
      with key matnr = data-matnr .

    if sy-subrc eq 0 .

      line-matkl = mara_line-matkl.
      line-meins = mara_line-meins.

    else .

      exit .

    endif .

    read table me->makt into makt_line
      with key matnr = data-matnr .

    if sy-subrc eq 0 .

      line-maktx = makt_line-maktx .

    else .

      concatenate data-matnr
                  space
                  '(NF'
                  space
                  data-indice
                  ')'
             into message .

      me->bapiret(
        exporting
          type       = 'E'
          message_v1 = 'Falha ao buscar informação'
          message_v2 = 'do material(MAKT)'
          message_v3 = message
        changing
          return     = return
      ).

      exit .

    endif .

    read table me->marc into marc_line
      with key matnr = data-matnr
               werks = data-werks .

    if sy-subrc eq 0 .

      line-nbm = marc_line-steuc .

    else .

      concatenate data-matnr
            space
            '(NF'
            space
      data-indice
            ')'
            into message .

      me->bapiret(
      exporting
        type       = 'E'
        message_v1 = 'Falha ao buscar informação'
        message_v2 = 'do material(MARC)'
        message_v3 = message
      changing
        return     = return
        ).

      exit .

    endif .


    read table me->mbew into mbew_line
      with key matnr = data-matnr
               bwkey = data-werks .

    if sy-subrc eq 0 .
      line-matuse = mbew_line-mtuse .
      line-matorg = mbew_line-mtorg .
    else .

      concatenate data-matnr
            space
            '(NF'
            space
      data-indice
            ')'
            into message .

      me->bapiret(
        exporting
          type       = 'E'
          message_v1 = 'Falha ao buscar informação'
          message_v2 = 'do material(MBEW)'
          message_v3 = message
        changing
          return     = return
          ).

      exit .

    endif .

    line-werks   = data-werks .
    line-charg   = data-charg .
    line-menge   = data-menge .
    line-cfop    = data-cfop .
    line-bwkey   = data-werks.
    line-docref  = data-docref.
    line-itmref  = data-itmref.
    line-menge   = data-menge.
    line-netpr   = data-netpr.
    line-netwr   = ( data-menge * data-netpr ).
    line-taxlw1  = data-taxlw1.
    line-taxlw2  = data-taxlw2.
    line-itmtyp  = '01' .
    line-incltx  = 'X' .
    line-werks   = data-werks .
    line-cfop_10 = data-cfop .
    line-taxlw4  = 'C49' .
    line-taxlw5  = 'P49' .

    append line to j_1bnflin .
    clear  line .

  endmethod .                    "get_j_1bnflin


  method get_j_1bnfstx .

    data:
      line      type bapi_j_1bnfstx,
      mara_line type me->ty_mara,
      message   type char50 .

    line-itmnum  = data-itmnum .
    line-taxtyp  = data-taxtyp .
    line-base    = data-base .
    line-rate    = data-rate .
    try .
        line-taxval  = ( data-base ) * ( data-rate / 100 ).
      catch cx_sy_zerodivide .
    endtry .
    line-excbas = data-excbas .

    append line to j_1bnfstx .
    clear  line .

    read table me->mara into mara_line
      with key matnr = data-matnr .

    if sy-subrc eq 0 .


      if mara_line-plgtp eq '00' .

*       Linha IPIS
        line-itmnum = data-itmnum .
        line-taxtyp = 'IPIS'.
        line-base   =  data-bspiscof .
        line-rate   = '1.65'.
        line-taxval = ( data-bspiscof ) * ( line-rate  / 100 ) .
        append line to j_1bnfstx .
        clear  line .

*       Linha ICOF
        line-itmnum = data-itmnum.
        line-taxtyp = 'ICOF '.
        line-base   = data-bspiscof.
        line-rate   = '7.60'.
        line-taxval = ( data-bspiscof ) * ( line-rate / 100 ).
        append line to j_1bnfstx .
        clear  line .

      endif.

    else .

      concatenate data-matnr
                  space
                  '(NF'
                  space
                  data-indice
                  ')'
             into message .

      me->bapiret(
        exporting
          type       = 'E'
          message_v1 = 'Falha ao buscar informação'
          message_v2 = 'do material(MARA)'
          message_v3 = message
        changing
          return     = return
      ).
    endif .

  endmethod .                    "get_j_1bnfstx

  method bapiret .

    data:
      line type bapiret2 .

    line-type       = type .
    line-id         = '>0' .
    line-number     = 000 .
    line-message_v1 = message_v1 .
    line-message_v2 = message_v2 .
    line-message_v3 = message_v3 .
    line-message_v4 = message_v4 .

    append line to return .
    clear  line .

  endmethod .                    "bapiret

  method get_zeile .

    data:
      mseg     type me->tab_mseg,
      line     type me->ty_mseg,
      out_line type me->ty_out .

    field-symbols:
      <data> type me->ty_out .

    loop at data assigning <data> .

      read table me->out into out_line
        with key matnr  = <data>-matnr
                 indice = <data>-indice .
      if sy-subrc eq 0 .
        <data>-mblnr = out_line-mblnr .
        <data>-mjahr = out_line-mjahr .
      endif .

    endloop .

*   Esperando 5 segundos ate o commit do banco de dados
    do 5 times .

      select mblnr mjahr zeile matnr
        from mseg
        into table mseg
         for all entries in data
       where mblnr eq data-mblnr
         and mjahr eq data-mjahr .

      if sy-subrc eq 0 .

        loop at mseg into line .

          read table data assigning <data>
            with key matnr = line-matnr .
          if sy-subrc eq 0 .
            <data>-zeile = line-zeile .
          endif .

        endloop .

        exit .
      else .
        wait up to 1 seconds .
      endif .

    enddo .

  endmethod .                    "get_zeile


  method get_percent .

    data:
      posicao type i .

*   posicao = me->convert_num( indice ) .
    posicao = indice .

    try .

        percent = ( 100 / total ) * indice .

      catch cx_sy_zerodivide .

    endtry .

    if percent ge 100 .
      percent = 95 .
    endif .


  endmethod .                    "get_percent

endclass .                    "lcl_class IMPLEMENTATION

*----------------------------------------------------------------------*
*- Varaveis locais
*----------------------------------------------------------------------*
data:
  obj type ref to lcl_class .

*----------------------------------------------------------------------*
*- Tela de seleção
*----------------------------------------------------------------------*
selection-screen begin of block b1 with frame title text-t01.

parameters:
  p_bukrs type bukrs     obligatory
          default '1000',
  p_lgort type lgort_d,
  p_file  type localfile obligatory
          default 'D:\Desktop\Temp\Panpharma\Carga BOL.xls' .

selection-screen end of block b1.

*----------------------------------------------------------------------*
*- Eventos
*----------------------------------------------------------------------*

initialization .

at selection-screen on value-request for p_file.
  lcl_class=>file_open_dialog(
    changing
      file = p_file
  ) .


start-of-selection .

  create object obj .

  obj->get_data(
    exporting
      bukrs = p_bukrs
      lgort = p_lgort
      file  = p_file
    ) .

  obj->show_data( ) .
