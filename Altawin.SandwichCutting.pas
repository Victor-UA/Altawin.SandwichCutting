Uses 'Victors', 'ProgressBar';
Const
  CurrDate=Now;

Var
  SQL: String;
  i: Integer;
  n: Integer;
  k: Integer;
  SCount: Integer;
  vFile: Text;
  FStream: TFileStream;
  FileName: String;
  Buffer: String;
  SandDL: icmDictionaryList;
  GlOrdersDL: icmDictionaryList;
  ErrOrdersDL: icmDictionaryList;
  GlOrder: IdocGlassOrder;
  GlOrderNames: String='';
  GlOrderTitle: String;
  GlTask: IdocGlassTask;
  GlTaskKey: Variant=-100500;
  GlTaskCapacity: Integer=0;
  Customer: idocCustomer;
  Owner: idocEmployee;
  Err: String;
  PInfo: String;
  pInfoBegin: String;
  pInfoModify: String;
  pInfoEnd: String;
  tmp: String;
  XMLDef: String;
  PBegin, PEnd: Integer;
  ArtGlassD: icmDictionary;
  GrOrderKeys: String;
  GUID: String;

procedure OnCloseProgram;
begin
  try
    PBDestroy;
    GLOrdersDL.Clear;
    ErrOrdersDL.Clear;
    SandDL.Clear;
  except
//    Showmessage('Помилка під час завершення програми'+#13+ExceptionMessage);
  end;
end;

Begin
  PBCreate;
  Application.ProcessMessages;

  //---------------------------------------------------------------------------- Завантаження XML замовлення склопакету
  try
    SQL:='select'+#13+
         '  v.var_bin'+#13+
         'from variables v'+#13+
         'where v.varname=''GLASSORDERDERPACKINFO''';
    XMLDef:=QueryValue(SQL, MakeDictionary([]));
  except
    showmessage('Не знайдено стандартну XML конструкції'+#13+ExceptionMessage+#13+SQL);
    OnCloseProgram;
    exit;
  end;

  GLOrdersDL:=CreateDictionaryList;

    //-------------------------------------------------------------------------- Визначення списку id кроїв МПК
  i:=0;
  while i<Length(Documents) do begin
    GrOrderKeys:=GrOrderKeys+VarToStr(Documents[i].Key)+',';
    inc(i);
  end;
  GrOrderKeys:=LeftStr(GrOrderKeys,Length(GrOrderKeys)-1);

  //---------------------------------------------------------------------------- Перевірка замовлення на присутність в іншому крої
  SQL:='select'+#13+
       '  go1.name,'+#13+
       '  list (distinct o.orderno, '', '') Orders'+#13+
       'from grordersdetail god'+#13+
       '  join orderitems oi on oi.orderitemsid=god.orderitemsid'+#13+
       '  join orders o on o.orderid=oi.orderid'+#13+
       '  join order_uf_values ov on ov.orderid=o.orderid and ov.userfieldid=(select uf.userfieldid from userfields uf where uf.fieldname=''grordersGLid'' and uf.doctype=''IdocWindowOrder'')'+#13+
       '  join grorders go1 on go1.guidhi=ov.var_guidhi and go1.guidlo=ov.var_guidlo'+#13+
       'where god.isaddition=0'+#13+
       '  and god.grorderid in ('+GrOrderKeys+')'+#13+
       'group by 1';
  try
    ErrOrdersDL:=QueryRecordList(SQL, MakeDictionary([]));
  except
    showmessage('Помилка отримання списку помилкових замовлень'+#13+ExceptionMessage+#13+SQL);
    OnCloseProgram;
    exit;
  end;
  try
    if ErrOrdersDL.Count>0 then begin
      Err:='Помилка створення крою: замовлення вже містять інші крої.'+#13;
      i:=0;
      while i<ErrOrdersDL.Count do begin
        Err:=Err+'Крой: ['+ErrOrdersDL[i]['name']+'], замовлення: ['+ErrOrdersDL[i]['Orders']+']'+#13;
        inc(i);
      end;
      Showmessage(Err);
      OnCloseProgram;
      exit;
    end;
  except
    showmessage('Системна помилка [замовлення вже містять інші крої]'+#13+ExceptionMessage);
    OnCloseProgram;
    exit;
  end;

  //---------------------------------------------------------------------------- Завантаження сендвічів, ініціалізація id структури WindowGlassOrder
  SQL:='select'+#13+
       '  gg.thick,'+#13+
       '  gg.marking,'+#13+
       '  go.name goname,'+#13+
       '  o.orderno,'+#13+
       '  oi.name as oiname,'+#13+
       '  gg.name,'+#13+
       '  gg.grgoodsid,'+#13+
       '  g.goodsid,'+#13+
       '  g.price1,'+#13+
       '  gpt.gptypeid,'+#13+
       '  gpt.marking gptMarking,'+#13+
       '  gpt.name gptName,'+#13+
       '  ev.valuesid evaluesid,'+#13+
       '  itd.width,'+#13+
       '  itd.height,'+#13+
       '  itd.qty*oi.qty qty,'+#13+
       '  itd.width*itd.height*itd.qty*oi.qty area,'+#13+
       '  itd.itemsdetailid,'+#13+
       '  iif('+#13+
       '    (select count(*) as qty'+#13+
       '    from orders o'+#13+
       '    where o.orderno=go.name'+#13+
       '    )=0,0,1'+#13+
       '  ) OrderExists,'+#13+
       '( select'+#13+
       '    gen_id(GEN_ORDERITEMS, 1)'+#13+
       '    from rdb$database'+#13+
       '  ) orderitemsid,'+#13+
       '( select'+#13+
       '    gen_id(GEN_MODELS, 1)'+#13+
       '    from rdb$database'+#13+
       '  ) modelid,'+#13+
       '( select'+#13+
       '    gen_id(GEN_MODELPARTS, 1)'+#13+
       '    from rdb$database'+#13+
       '  ) modelpartid,'+#13+
       '( select'+#13+
       '    gen_id(GEN_ITEMSDETAIL, 1)'+#13+
       '    from rdb$database'+#13+
       '  ) NewItemsDetailId,'+#13+
       '( select'+#13+
       '    gen_id(GEN_GRORDERSDETAIL, 1)'+#13+
       '    from rdb$database'+#13+
       '  ) grorderdetailid,'+#13+
       ' go.grorderid,'+#13+
       ' 0 as GlOrderId,'+#13+
       '  '''' as thumbs,'+#13+
       ' -1 as GlOrdersDLIndex'+#13+
       'from grordersdetail god'+#13+
       '  join grorders go on go.grorderid=god.grorderid'+#13+
       '  join orderitems oi on oi.orderitemsid=god.orderitemsid'+#13+
       '  join orders o on o.orderid=oi.orderid'+#13+
       '  join itemsdetail itd on itd.orderitemsid=oi.orderitemsid'+#13+
       '  join groupgoods gg on gg.grgoodsid=itd.grgoodsid'+#13+
       '  join goods g on g.goodsid=itd.goodsid'+#13+
       '  join groupgoodstypes ggt on ggt.ggtypeid=gg.ggtypeid'+#13+
       '  left join e_values ev on ev.grgoodsid=gg.grgoodsid'+#13+
       '  join e_valuestree evt on evt.valuestreeid=ev.parentid'+#13+
       '  join e_valuestree evtp on evtp.valuestreeid=evt.parentid and evtp.nodetitle=''Сэндвич'''+#13+
       '  join gpackettypes gpt on evt.noderule containing '''''''' || gpt.marking || '''''''''+#13+
       'where ggt.code=''Sendvich'''+#13+
       '  and god.isaddition=0'+#13+
       '  and go.grorderid in ('+GrOrderKeys+')'+#13+
       'order by 1,2,3,4,5';
  try
//    showmessage(SQL);//!
    SandDL:=QueryRecordList(SQL, MakeDictionary([]));
    fPB.Max:=SandDL.Count;
    OldP:=0;
  except
    fPB.Max:=0;
    showmessage('Помилка запиту сендвічів'+#13+ExceptionMessage+#13+SQL);
    OnCloseProgram;
    exit;
  end;

  //---------------------------------------------------------------------------- Ініціалізація id структури кроїв склопакетів
  if SandDL.Count>0 then begin
    SQL:='select'+#13+
         '  go.grorderid,'+#13+
         '  go.name,'+#13+
         '  go.groupdate,'+#13+
         '  ( select'+#13+
         '      gen_id(GEN_ORDERS,1)'+#13+
         '    from rdb$database'+#13+
         '  ) as GlOrderId,'+#13+
         '  ( select'+#13+
         '      gen_id(GEN_APPROVEDOCUMENTS,1)'+#13+
         '    from rdb$database'+#13+
         '  ) as approvedocumentid,'+#13+
         '  ( select'+#13+
         '      cu.customerid'+#13+
         '    from customers cu'+#13+
         '      join contragents ca on ca.contragid=cu.contragid'+#13+
         '    where ca.name=''Газда'''+#13+
         '  ) as customerid,'+#13+
         '  0 as isCreated'+#13+
         'from grorders go'+#13+
         'where go.grorderid in ('+GrOrderKeys+')';
    try
      GlOrdersDL:=QueryRecordList(SQL, MakeDictionary([]));
    except
      showmessage('Помилка визначення GlOrderId'+#13+ExceptionMessage+#13+SQL);
      OnCloseProgram;
      exit;
    end;
  end;

  n:=0;
  while n<SandDL.Count do begin
    i:=0;
    while (i<GlOrdersDL.Count) and (GlOrdersDL[i]['grorderid']<>SandDL[n]['grorderid']) do begin
      inc(i);
    end;
    if GlOrdersDL[i]['grorderid']<>SandDL[n]['grorderid'] then begin
      RaiseException('Не знайдено відповідного крою:'+#13+'GlOrdersDL['+VarToStr(i)+'].Key='+VarToStr(GlOrdersDL[i]['grorderid'])+'<>'+'SandDL['+VarToStr(n)+'][''grorderid'']');
    end;
    SandDL[n]['GlOrderId']:=GlOrdersDL[i]['GlOrderId'];
    SandDL[n]['GlOrdersDLIndex']:=i;

    if GlTaskKey=-100500 then begin
      SQL:='select'+#13+
           '  gen_id(GEN_GRORDERS, 1)'+#13+
           'from rdb$database';
      try
        GlTaskKey:=QueryValue(SQL, MakeDictionary([]));
      except
        showmessage('Помилка визначення GrOrderId'+#13+ExceptionMessage+#13+SQL);
        OnCloseProgram;
        exit;
      end;

      //------------------------------------------------------------------------ Створення крою склопакетів
      SQL:='insert into grorders (grorderid, name, isoptimized, isdefault, makebill, rcomment, reccolor, recflag, guidhi, guidlo, ownerid, datecreated, datemodified, datedeleted, ownertype, procschemaid, groupdate, capacity, productionsid, planid, isclosed, deleted, linearoptim, linearopttype, linearuserest, linearsaverest, linearrestmode, linearpaired, layoutoptim, layoutopttype, layoutuserest, layoutsaverest, whlistid, ordernames)'+#13+
           'values ('+#13+
           '  :grorderid,'+#13+
           '  :name,'+#13+
           '  0,'+#13+
           '  0,'+#13+
           '  1,'+#13+
           '  :rcomment,'+#13+
           '  null,'+#13+
           '  null,'+#13+
           '  :guidhi,'+#13+
           '  :guidlo,'+#13+
           '  :ownerid,'+#13+
           '  :datecreated,'+#13+
           '  :datemodified,'+#13+
           '  null,'+#13+
           '  1,'+#13+
           '  null,'+#13+
           '  :groupdate,'+#13+
           '  :capacity,'+#13+       //Заповнити після виконання циклу
           '  null,'+#13+
           '  null,'+#13+
           '  0,'+#13+
           '  0,'+#13+
           '  0,'+#13+
           '  0,'+#13+
           '  0,'+#13+
           '  0,'+#13+
           '  null,'+#13+
           '  0,'+#13+
           '  1,'+#13+
           '  1,'+#13+
           '  1,'+#13+
           '  1,'+#13+
           '  0,'+#13+
           '  :ordernames'+#13+      //Заповнити після виконання циклу
           ')';
      GUID:=GenerateGUID;
      SQL:=ReplaceText(SQL, ':grorderid', VarToSQL(GlTaskKey));
      SQL:=ReplaceText(SQL, ':name', VarToSQL('Сендвіч від '+DateTimeToStr(CurrDate)));
      SQL:=ReplaceText(SQL, ':rcomment', VarToSQL('Створено автоматично'));
      SQL:=ReplaceText(SQL, ':guidhi', VarToSQL(GUIDHi(GUID)));
      SQL:=ReplaceText(SQL, ':guidlo', VarToSQL(GUIDLo(GUID)));
      SQL:=ReplaceText(SQl, ':ownerid', VarToSQL(UserContext.UserID));
      SQL:=ReplaceText(SQl, ':datecreated', VarToSQL(Now));
      SQL:=ReplaceText(SQl, ':datemodified', VarToSQL(Now));
      SQL:=ReplaceText(SQL, ':groupdate', VarToSQL(Now));
      SQL:=ReplaceText(SQL, ':capacity', VarToSQL(0));
      SQL:=ReplaceText(SQL, ':ordernames', VarToSQL(''));
      try
        Err:=ExecuteSQLCommit(SQL);
      except
        showmessage('Помилка створення змінного завдання склопакетів'+#13+ExceptionMessage+#13+SQL);
        OnCloseProgram;
        exit;
      end;
    end;

    if GlOrdersDL[i]['isCreated']=0 then begin
      try
        GlOrderTitle:=GlOrdersDL[i]['name'];
        if SandDL[n]['OrderExists']>0 then begin
          GlOrderTitle:=GlOrderTitle+' ('+DateTimeToStr(CurrDate)+')';
        end;
        SQL:='insert into approvedocuments (approvedocumentid, doctype)'+#13+
             'values ('+#13+
             '  :approvedocumentid,'+#13+
             '  ''IdocOrder'''+#13+
             ')';
        SQL:=ReplaceText(SQL, ':approvedocumentid', VarToSQL(GlOrdersDL[i]['approvedocumentid']));
        Err:=ExecuteSQLCommit(SQL);
        if Err<>'' then begin
          showmessage('Помилка створення [approvedocuments]'+#13+Err);
          OnCloseProgram;
          exit;
        end;

        //---------------------------------------------------------------------- Створення замовлення склопакетів
        SQL:='insert into orders (orderid, ownertype, orderno, agreementno, agreementdate, currencyid, sellerid, customerid, itemstatusmode, totalpricelock, proddate, dateorder, orderstatus, lastgenitem, guidhi, guidlo, ownerid, datecreated, datemodified, deleted, rcomment, valid, totalprice, payment, isdealeradd, isdealerstartadd, isreserved, approvedocumentid, crossrate)'+#13+
             'values ('+#13+
             '  :orderid,'+#13+
             '  1,'+#13+
             '  :orderno,'+#13+
             '  null,'+#13+
             '  null,'+#13+
             '  null,'+#13+
             '  :sellerid,'+#13+
             '  :customerid,'+#13+
             '  0,'+#13+
             '  0,'+#13+
             '  :proddate,'+#13+
             '  :dateorder,'+#13+
             '  0,'+#13+
             '  0,'+#13+
             '  :guidhi,'+#13+
             '  :guidlo,'+#13+
             '  :ownerid,'+#13+
             '  :datecreated,'+#13+
             '  :datemodified,'+#13+
             '  0,'+#13+
             '  :rcomment,'+#13+
             '  1,'+#13+
             '  0,'+#13+
             '  0,'+#13+
             '  0,'+#13+
             '  0,'+#13+
             '  0,'+#13+
             '  :approvedocumentid,'+#13+
             '  1'+#13+
             ')';
        GUID:=GenerateGUID;
        SQL:=ReplaceText(SQl, ':orderid', VarToSQL(GlOrdersDL[i]['GlOrderId']));
        SQL:=ReplaceText(SQl, ':orderno', VarToSQL(GlOrderTitle));
        SQL:=ReplaceText(SQl, ':sellerid', VarToSQL('null'));
        SQL:=ReplaceText(SQl, ':customerid', VarToSQL(GlOrdersDL[i]['customerid']));
        SQL:=ReplaceText(SQl, ':proddate', VarToSQL(GlOrdersDL[i]['groupdate']));
        SQL:=ReplaceText(SQl, ':dateorder', VarToSQL(GlOrdersDL[i]['groupdate']));
        SQL:=ReplaceText(SQl, ':guidhi', VarToSQL(GUIDHi(GUID)));
        SQL:=ReplaceText(SQl, ':guidlo', VarToSQL(GUIDLo(GUID)));
        SQL:=ReplaceText(SQl, ':ownerid', VarToSQL(UserContext.UserID));
        SQL:=ReplaceText(SQl, ':datecreated', VarToSQL(Now));
        SQL:=ReplaceText(SQl, ':datemodified', VarToSQL(Now));
        SQL:=ReplaceText(SQl, ':rcomment', VarToSQL('Вибірка сендвічів з крою ['+GlOrdersDL[i]['name']+']'));
        SQL:=ReplaceText(SQl, ':approvedocumentid', VarToSQL(GlOrdersDL[i]['approvedocumentid']));
        Err:=ExecuteSQLCommit(SQL);
        if Err<>'' then begin
          showmessage('Помилка створення замовлення склопакетів'+#13+Err);
          OnCloseProgram;
          exit;
        end;
        GlOrdersDL[i]['isCreated']:=1;
      except
        showmessage('Помилка створення замовлення склопакетів'+#13+ExceptionMessage+#13+GlOrderTitle);
        OnCloseProgram;
        exit;
      end;
    end;

    //-------------------------------------------------------------------------- Створення XML моделі склопакету
    PInfo:=XMLDef;
    pBegin:=-1;
    PInfo:=SetXMLParamValue(PInfo, 'ARTGLASSID', 'Any value', SandDL[n]['gptMarking']);
    PInfo:=SetXMLParamValue(PInfo, 'ARTGLASSART', 'Any value', SandDL[n]['gptName']);
    pInfoBegin:='';
    pInfoEnd:=PInfo;
    k:=1;
    while (k<=4) do begin
      case k of
        1: begin
             pInfoEnd:=SetXMLParamValue(pInfoEnd, 'X1', 'any value', VarToStr(SandDL[n]['width']), -1, 1);
             pInfoEnd:=SetXMLParamValue(pInfoEnd, 'Y1', 'any value', VarToStr(SandDL[n]['height']), -1, 1);
             pInfoEnd:=SetXMLParamValue(pInfoEnd, 'X2', 'any value', '0', -1, 1);
             pInfoEnd:=SetXMLParamValue(pInfoEnd, 'Y2', 'any value', VarToStr(SandDL[n]['height']), -1, 1);
           end;
        2: begin
             pInfoEnd:=SetXMLParamValue(pInfoEnd, 'X1', 'any value', '0', -1, 1);
             pInfoEnd:=SetXMLParamValue(pInfoEnd, 'Y1', 'any value', VarToStr(SandDL[n]['height']), -1, 1);
             pInfoEnd:=SetXMLParamValue(pInfoEnd, 'X2', 'any value', '0', -1, 1);
             pInfoEnd:=SetXMLParamValue(pInfoEnd, 'Y2', 'any value', '0', -1, 1);
           end;
        3: begin
             pInfoEnd:=SetXMLParamValue(pInfoEnd, 'X1', 'any value', '0', -1, 1);
             pInfoEnd:=SetXMLParamValue(pInfoEnd, 'Y1', 'any value', '0', -1, 1);
             pInfoEnd:=SetXMLParamValue(pInfoEnd, 'X2', 'any value', VarToStr(SandDL[n]['width']), -1, 1);
             pInfoEnd:=SetXMLParamValue(pInfoEnd, 'Y2', 'any value', '0', -1, 1);
           end;
        4: begin
             pInfoEnd:=SetXMLParamValue(pInfoEnd, 'X1', 'any value', VarToStr(SandDL[n]['width']), -1, 1);
             pInfoEnd:=SetXMLParamValue(pInfoEnd, 'Y1', 'any value', '0', -1, 1);
             pInfoEnd:=SetXMLParamValue(pInfoEnd, 'X2', 'any value', VarToStr(SandDL[n]['width']), -1, 1);
             pInfoEnd:=SetXMLParamValue(pInfoEnd, 'Y2', 'any value', VarToStr(SandDL[n]['height']), -1, 1);
           end;
      end;
      pInfoModify:=copy(pInfoEnd, 0, Pos('</BALKA>',pInfoEnd)+length('</BALKA>')-1);
      pInfoBegin:=pInfoBegin+pInfoModify;
      pInfoEnd:=copy(pInfoEnd, Pos('</BALKA>',pInfoEnd)+length('</BALKA>'), Length(pInfoEnd)-Length(pInfoModify));
      inc(k);
    end;
    PInfo:=pInfoBegin+pInfoEnd;

    //-------------------------------------------------------------------------- Спроба генерувати зображення склопакету
{
    try
      FileName:='C:\TEMP\'+VarToStr(SandDL[n]['OrderItemsId'])+'.png';
      CreateModelImageFile(PInfo, 0, FileName, SandDL[n]['width'], SandDL[n]['height']);
      if FileExists(FileName) then begin
        try
          FStream:=TFileStream.Create(FileName, fmRead);
          FStream.Seek(0,0);
          FStream.Read(SandDL[n]['thumbs'], FStream.Size);
          FStream.Free;
          Execute('del '+FileName+' /Q');
        except
          showmessage('Помилка завантаження зображення'+#13+ExceptionMessage);
        end;
      end;
    except
      showmessage('Помилка генерування зображення конструкції'+#13+ExceptionMessage);
    end;
}

    //-------------------------------------------------------------------------- Створення конструкції склопакету
    SQL:='insert into VTORDERITEMSGL (orderitemsid, orderid, name, qty, laboriousness, area, rcomment, isaddition, usedqty, usedaddqty, thumbs, valid, price, costall, packinfo, productcount, width, height, GPTYPEID, FORMULA, GEOMETRY, SHPROSSES)'+#13+
         'values ('+#13+
         '  :orderitemsid,'+#13+
         '  :orderid,'+#13+
         '  :name,'+#13+
         '  :qty,'+#13+
         '  1,'+#13+
         '  :area,'+#13+
         '  :rcomment,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  :thumbs,'+#13+
         '  1,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  :packinfo,'+#13+
         '  1,'+#13+
         '  :width,'+#13+
         '  :height,'+#13+
         '  :GPTYPEID,'+#13+
         '  :FORMULA,'+#13+
         '  :GEOMETRY,'+#13+
         '  :SHPROSSES'+#13+
         ')';
    SQL:=ReplaceText(SQL, ':orderitemsid', VarToSQL(SandDL[n]['OrderItemsId']));
    SQL:=ReplaceText(SQL, ':orderid', VarToSQL(SandDL[n]['GlOrderId']));
    SQL:=ReplaceText(SQL, ':name', VarToSQL(PADL(VarToStr(n),2,'0')+'['+copy(SandDL[n]['orderno']+'-'+SandDL[n]['oiname'],1,28)+']'));
    SQL:=ReplaceText(SQL, ':qty', VarToSQL(SandDL[n]['qty']));
    SQL:=ReplaceText(SQL, ':area', VarToSQL(SandDL[n]['area']));
    SQL:=ReplaceText(SQL, ':rcomment', VarToSQL(SandDL[n]['marking']+' '+SandDL[n]['name']+' ['+VarToStr(SandDL[n]['itemsdetailid'])+']-['+SandDL[n]['oiname']+']/['+SandDL[n]['orderno']+']'));
    SQL:=ReplaceText(SQL, ':thumbs', VarToSQL('null'));
//    SQL:=ReplaceText(SQL, ':thumbs', VarToSQL(SandDL[n]['thumbs']));
    SQL:=ReplaceText(SQL, ':packinfo', VarToSQL(PInfo));
    SQL:=ReplaceText(SQL, ':width', VarToSQL(SandDL[n]['width']));
    SQL:=ReplaceText(SQL, ':height', VarToSQL(SandDL[n]['height']));
    SQL:=ReplaceText(SQL, ':GPTYPEID', VarToSQL(SandDL[n]['gptypeid']));
    SQL:=ReplaceText(SQL, ':FORMULA', VarToSQL(SandDL[n]['gptName']));
    SQL:=ReplaceText(SQL, ':GEOMETRY', VarToSQL(0));
    SQL:=ReplaceText(SQL, ':SHPROSSES', VarToSQL(0));
//    Err:=ExecuteSQLCommit(SQL, '', MakeDictionary(['thumbs', SandDL[n]['thumbs']]));
    Err:=ExecuteSQLCommit(SQL);
    if Err<>'' then begin
      showmessage('Помилка створення конструкції'+#13+Err);
      OnCloseProgram;
      exit;
    end;
{
    SQL:='update orderitems set thumbs=:thumbs where orderitemsid='+VarToSQL(SandDL[n]['OrderItemsId']);
    try
      ExecSQL(SQL, MakeDictionary(['thumbs', SandDL[n]['thumbs']]));
      showmessage(1);//!
    except
      showmessage('Помилка запису зображення конструкції'+#13+ExceptionMessage+#13+SQL);
      OnCloseProgram;
      exit;
    end;
}

    //-------------------------------------------------------------------------- Створення моделі склопакету
    SQL:='insert into models (modelid, orderitemsid, modelno, modelwidth, modelheight, flugelcount, area, incolorid, outcolorid, modelthick)'+#13+
         'values ('+#13+
         '  :modelid,'+#13+
         '  :orderitemsid,'+#13+
         '  0,'+#13+
         '  :modelwidth,'+#13+
         '  :modelheight,'+#13+
         '  0,'+#13+
         '  :area,'+#13+
         '  1,'+#13+
         '  1,'+#13+
         '  :modelthick'+#13+
         ')';
    SQL:=ReplaceText(SQL, ':modelid', VarToSQL(SandDL[n]['ModelId']));
    SQL:=ReplaceText(SQL, ':orderitemsid', VarToSQL(SandDL[n]['OrderItemsId']));
    SQL:=ReplaceText(SQL, ':modelwidth', VarToSQL(SandDL[n]['width']));
    SQL:=ReplaceText(SQL, ':modelheight', VarToSQL(SandDL[n]['height']));
    SQL:=ReplaceText(SQL, ':area', VarToSQL(SandDL[n]['width']*SandDL[n]['height']));
    SQL:=ReplaceText(SQL, ':modelthick', VarToSQL(SandDL[n]['thick']));
    Err:=ExecuteSQLCommit(SQL);
    if Err<>'' then begin
      showmessage('Помилка створення моделі'+#13+Err);
      OnCloseProgram;
      exit;
    end;

    //-------------------------------------------------------------------------- Створення частин моделі склопакету
    SQL:='insert into modelparts (modelpartid, sprpartid, partnum, flugelopentype, flugelopentag, ishandle, handlepos, handleposfalts, modelid, partwidth, partheight, gptypeid, elementuid, partthick)'+#13+
         'values ('+#13+
         '  :modelpartid,'+#13+
         '  3,'+#13+
         '  1,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  :modelid,'+#13+
         '  :partwidth,'+#13+
         '  :partheight,'+#13+
         '  :gptypeid,'+#13+
         '  :elementuid,'+#13+
         '  :partthick'+#13+
         ')';
    SQL:=ReplaceText(SQL, ':modelpartid', VarToSQL(SandDL[n]['ModelPartId']));
    SQL:=ReplaceText(SQL, ':modelid', VarToSQL(SandDL[n]['ModelId']));
    SQL:=ReplaceText(SQL, ':partwidth', VarToSQL(SandDL[n]['width']));
    SQL:=ReplaceText(SQL, ':partheight', VarToSQL(SandDL[n]['height']));
    SQL:=ReplaceText(SQL, ':gptypeid', VarToSQL(SandDL[n]['gptypeid']));
    SQL:=ReplaceText(SQL, ':elementuid', VarToSQL('{1693328F-49C4-4252-B96C-75D7DF9F4A31}'));
    SQL:=ReplaceText(SQL, ':partthick', VarToSQL(SandDL[n]['thick']));
    Err:=ExecuteSQLCommit(SQL);
    if Err<>'' then begin
      showmessage('Помилка створення частини моделі'+#13+Err);
      OnCloseProgram;
      exit;
    end;

    //-------------------------------------------------------------------------- Створення деталізації конструкції  склопакету
    SQL:='insert into itemsdetail (itemsdetailid,orderitemsid,setindex,grgoodsid,goodsid,modelpartid,positionid,partnum,modelno,width,height,thick,qty,ang1,ang2,radius,pricetype,weight,connection1,connection2,isextended,updatestatus,rcomment,allvolume,allsavingvolume,allweight,price,savingabs,cost,savingcost,int_marking,izdpart,partside,evaluesid,mark,elementuid)'+#13+
         'values ('+#13+
         '  :itemsdetailid,'+#13+
         '  :orderitemsid,'+#13+
         '  0,'+#13+
         '  :grgoodsid,'+#13+
         '  :goodsid,'+#13+
         '  :modelpartid,'+#13+
         '  0,'+#13+
         '  1,'+#13+
         '  0,'+#13+
         '  :width,'+#13+
         '  :height,'+#13+
         '  1,'+#13+
         '  :qty,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  :rcomment,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  :price,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  0,'+#13+
         '  :int_marking,'+#13+
         '  ''Г-1'','+#13+
         '  '''','+#13+
         '  :evaluesid,'+#13+
         '  '''','+#13+
         '  :elementuid'+#13+
         ')';
    SQL:=ReplaceText(SQL, ':itemsdetailid', VarToSQL(SandDL[n]['NewItemsDetailId']));
    SQL:=ReplaceText(SQL, ':orderitemsid', VarToSQL(SandDL[n]['orderitemsid']));
    SQL:=ReplaceText(SQL, ':grgoodsid', VarToSQL(SandDL[n]['grgoodsid']));
    SQL:=ReplaceText(SQL, ':goodsid', VarToSQL(SandDL[n]['goodsid']));
    SQL:=ReplaceText(SQL, ':modelpartid', VarToSQL(SandDL[n]['ModelPartId']));
    SQL:=ReplaceText(SQL, ':width', VarToSQL(SandDL[n]['width']));
    SQL:=ReplaceText(SQL, ':height', VarToSQL(SandDL[n]['height']));
    SQL:=ReplaceText(SQL, ':qty', VarToSQL(SandDL[n]['qty']));
    SQL:=ReplaceText(SQL, ':qty', VarToSQL(SandDL[n]['qty']));
    SQL:=ReplaceText(SQL, ':rcomment', VarToSQL('['+VarToStr(SandDL[n]['itemsdetailid'])+']-['+SandDL[n]['oiname']+']'));
    SQL:=ReplaceText(SQL, ':price', VarToSQL(SandDL[n]['price1']));
    SQL:=ReplaceText(SQL, ':int_marking', VarToSQL(SandDL[n]['gptMarking']));
    SQL:=ReplaceText(SQL, ':evaluesid', VarToSQL(SandDL[n]['evaluesid']));
    SQL:=ReplaceText(SQL, ':elementuid', VarToSQL('{1693328F-49C4-4252-B96C-75D7DF9F4A31}'));
    Err:=ExecuteSQLCommit(SQL);
    if Err<>'' then begin
      showmessage('Помилка створення [itemsdetail]'+#13+Err);
      OnCloseProgram;
      exit;
    end;

    //-------------------------------------------------------------------------- Створення деталізації крою склопакетів
    SQL:='insert into grordersdetail (grorderdetailid, grorderid, orderitemsid, qty, isaddition)'+#13+
         'values ('+#13+
         '  :grorderdetailid,'+#13+
         '  :grorderid,'+#13+
         '  :orderitemsid,'+#13+
         '  :qty,'+#13+
         '  :isaddition'+#13+
         ')';
    SQL:=ReplaceText(SQL, ':grorderdetailid', VarToSQL(SandDL[n]['grorderdetailid']));
    SQL:=ReplaceText(SQL, ':grorderid', VarToSQL(GlTaskKey));
    SQL:=ReplaceText(SQL, ':orderitemsid', VarToSQL(SandDL[n]['orderitemsid']));
    SQL:=ReplaceText(SQL, ':qty', VarToSQL(SandDL[n]['qty']));
//    SQL:=ReplaceText(SQL, ':qty', VarToSQL(1));                               2015.11.05
    SQL:=ReplaceText(SQL, ':isaddition', VarToSQL(0));
    Err:=ExecuteSQLCommit(SQL);
    if Err<>'' then begin
      showmessage('Помилка створення [grordersdetail]'+#13+Err);
      OnCloseProgram;
      exit;
    end;

    PBStep(SandDL[n]['marking']+'/'+SandDL[n]['goname']+'/'+SandDL[n]['orderno']+'/'+SandDL[n]['oiname']);

    GlTaskCapacity:=GlTaskCapacity+GlOrdersDL[i]['qty'];
    inc(n);
  end;

  FormPB.Hide;

  //---------------------------------------------------------------------------- Розрахунок крою склопакетів
  i:=0;
  while i<GlOrdersDL.Count do begin
    if GlOrdersDL[i]['isCreated']=1 then begin
      try
        GlOrder:=OpenDocument(IdocGlassOrder, GlOrdersDL[i]['GlOrderId']);
        GlOrder.Calculate;
        GlOrder.Save;

      except
        showmessage('Помилка калькуляції замовлення склопакетів ['+GlOrdersDL[i]['name']+']'+#13+ExceptionMessage);
        OnCloseProgram;
        exit;
      end;
    end;
    inc(i);
  end;
  GlOrder:=null;

  if GlTaskKey<>-100500 then begin
{
    while i<GlOrdersDL.Count do begin
      if GlOrdersDL[i]['isCreated']=1 then begin
        GlOrderNames:=GlOrderNames+GlOrdersDL[i]['name']+', ';
      end;
      inc(i);
    end;
    GlOrderNames:=Copy(GlOrderNames, 1, Length(GlOrderNames)-2);
}

    //-------------------------------------------------------------------------- Запис кількості склопакетів у крою склопакетів
    SQL:='update grorders go'+#13+
         'set'+#13+
         '  go.capacity=:capacity'+#13+
//         '  go.ordernames=:ordernames'+#13+
         'where go.grorderid=:grorderid';
    SQL:=ReplaceText(SQL, ':capacity', VarToSQL(GlTaskCapacity));
    SQL:=ReplaceText(SQL, ':ordernames', VarToSQL(GlOrderNames));
    SQL:=ReplaceText(SQL, ':grorderid', VarToSQL(GlTaskKey));
    try
      Err:=ExecuteSQLCommit(SQL);
    except
      showmessage('Помилка редагування змінного завдання склопакетів'+#13+ExceptionMessage+#13+SQL);
      OnCloseProgram;
      exit;
    end;

    //-------------------------------------------------------------------------- Прив'язка крою склопакетів у замовленні конструкції
    SQL:='execute block'+#13+
         'as'+#13+
         '  declare variable orderid id;'+#13+
         'Begin'+#13+
         '  for'+#13+
         '    select'+#13+
         '      o.orderid'+#13+
         '    from orders o'+#13+
         '    where o.ownertype=0'+#13+
         '      and exists ('+#13+
         '        select'+#13+
         '          oi.orderid'+#13+
         '        from grordersdetail gd'+#13+
         '          join orderitems oi on oi.orderitemsid=gd.orderitemsid'+#13+
         '        where gd.grorderid in (:grorders)'+#13+
         '          and oi.orderid=o.orderid'+#13+
         '      )'+#13+
         '  into :orderid'+#13+
         '  do begin'+#13+
         '    update or insert into order_uf_values (orderid, userfieldid, var_str, var_guidhi, var_guidlo)'+#13+
         '    values ('+#13+
         '      :orderid,'+#13+
         '      (select uf.userfieldid from userfields uf where uf.fieldname=''grordersGLid'' and uf.doctype=''IdocWindowOrder''),'+#13+
         '      (select go.name from grorders go where go.grorderid = :grorderid),'+#13+
         '      (select go.guidhi from grorders go where go.grorderid = :grorderid),'+#13+
         '      (select go.guidlo from grorders go where go.grorderid = :grorderid)'+#13+
         '    )'+#13+
         '    matching(orderid, userfieldid);'+#13+
         '  end'+#13+
         'End';
    SQL:=ReplaceText(SQL, ':grorders', GrOrderKeys);
    SQL:=ReplaceText(SQL, ':grorderid', VarToSQL(GlTaskKey));
    try
      Err:=ExecuteSQLCommit(SQL);
    except
      showmessage('Помилка прив''язки змінного завдання склопакетів до замовлення конструкцій'+#13+ExceptionMessage+#13+SQL);
      OnCloseProgram;
      exit;
    end;

    //-------------------------------------------------------------------------- Підрахунок кількості створених кроїв склопакетів
    n:=0;
    i:=0;
    while i<GlOrdersDL.Count do begin
      if GlOrdersDL[i]['isCreated']=1 then begin
        inc(n)
      end;
      inc(i);
    end;
    if UserContext.UserName='victor' then begin
      showmessage('Оброблено:'+#13+
                  '  Кроїв: '+VarToStr(GLOrdersDL.Count)+#13+
                  '  Кроїв з сендвічем: '+VarToStr(n)
      );
    end;
    GlTask:=OpenDocument(IdocGlassTask, GlTaskKey);
    GlTask.Show;
  end;
  OnCloseProgram;
End;
