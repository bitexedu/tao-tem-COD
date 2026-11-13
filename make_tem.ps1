param(
  [string]$root = 'C:\Users\thaohlt.BITEXDC0\Desktop\Trình đơn hàng',
  [string]$bangKe = '',
  [string]$baseTem = '',
  [string]$outName = '',
  [string]$report = ''
)
$ci=[Globalization.CultureInfo]::InvariantCulture
$temDir=Join-Path $root 'Tem COD'
$log=@()
if($bangKe -eq ''){ $bangKe=(Get-ChildItem $temDir -Filter 'BANG KE*.xlsx' | Where-Object { $_.Name -notlike '~*' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1).Name }
if($outName -eq ''){ $outName=($bangKe -replace '^BANG KE','TEM COD') -replace '\.xlsx$',' - moi.xlsx' }
$baseTemPath= if($baseTem -eq ''){ Join-Path $root '_auto\base\Tem_base.xlsx' } else { Join-Path $temDir $baseTem }
function L($s){ $script:log+=$s; $s }
function S($x){ if($null -eq $x){''}else{([string]$x).Trim()} }
function Phone($x){ $s=(S $x) -replace '\D',''; if($s.Length -eq 9){ '0'+$s }else{ $s } }
function ReadSheet($wb,$sheet){ $ws=$wb.Worksheets.Item($sheet); $ur=$ws.UsedRange; [pscustomobject]@{r0=$ur.Row;c0=$ur.Column;v=$ur.Value2;nr=$ur.Rows.Count;nc=$ur.Columns.Count} }
function Cell($t,$r,$c){ $i=$r-$t.r0+1; $j=$c-$t.c0+1; if($i -lt 1 -or $j -lt 1 -or $i -gt $t.nr -or $j -gt $t.nc){return $null}; $t.v[$i,$j] }

$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
try{
# ---------- 1. BẢNG KÊ ----------
$bkPath=Join-Path $temDir $bangKe
$wbK=$xl.Workbooks.Open($bkPath)
$k=ReadSheet $wbK 1
$hdr=@{}; for($c=1;$c -le $k.nc;$c++){ $hdr[(S (Cell $k 1 $c))]=$c }
$cPD=$hdr['Posting Date']; $cInv=$hdr['NWV VAT Number.']; $cNo=$hdr['No.']; $cDesc=$hdr['Description']; $cAmt=$hdr['Line Amount Incl. VAT']; $cQty=$hdr['Quantity (Base)']; $cWD=$hdr['Work Description']
$lines=@()
for($r=2;$r -le $k.r0+$k.nr-1;$r++){
  $inv=S (Cell $k $r $cInv); if($inv -eq ''){continue}
  $lines+=[pscustomobject]@{pd=[double](Cell $k $r $cPD);inv=$inv;no=(S (Cell $k $r $cNo));amt=[double](Cell $k $r $cAmt);wd=(S (Cell $k $r $cWD))}
}
$dates=$lines | Select-Object -ExpandProperty pd -Unique | Sort-Object
L ("Bang ke: {0} dong, Posting Date: {1}" -f $lines.Count, (($dates | % { [datetime]::FromOADate($_).ToString('dd/MM/yyyy') }) -join ', '))

# ---------- 2. PIVOT ----------
$existing=$null; foreach($s in $wbK.Worksheets){ if($s.Name -eq 'Pivot'){ $existing=$s } }
if($existing){ $existing.Delete() }
$wsK=$wbK.Worksheets.Item(1)
$pc=$wbK.PivotCaches().Create(1,$wsK.UsedRange)
$pws=$wbK.Worksheets.Add([Type]::Missing,$wbK.Worksheets.Item($wbK.Worksheets.Count)); $pws.Name='Pivot'
$pt=$pc.CreatePivotTable($pws.Range('A3'),'PivotBangKe')
$pos=1
foreach($fn in 'Work Description','NWV VAT Number.','No.','Variant Code','Description'){
  $pf=$pt.PivotFields($fn); $pf.Orientation=1; $pf.Position=$pos; $pos++
}
$d1=$pt.AddDataField($pt.PivotFields('Quantity (Base)'),'Sum - Quantity (Base)',-4157); $d1.NumberFormat='#,##0'
$d2=$pt.AddDataField($pt.PivotFields('Line Amount Incl. VAT'),'Sum - Line Amount Incl. VAT',-4157); $d2.NumberFormat='#,##0'
$pt.RowAxisLayout(1)
$pt.RepeatAllLabels(2)
foreach($fn in 'Work Description','NWV VAT Number.','No.','Variant Code','Description'){
  $pf=$pt.PivotFields($fn)
  try{ $pf.Subtotals(1)=$false }catch{ try{ [void][System.__ComObject].InvokeMember('Subtotals',[Reflection.BindingFlags]::SetProperty,$null,$pf,@(1,$false)) }catch{} }
}
$pt.CompactLayoutRowHeader='Work Description'
$pws.Range('A1').Value2="PIVOT BẢNG KÊ - Posting Date $(($dates | % { [datetime]::FromOADate($_).ToString('dd/MM/yyyy') }) -join ', ')"
$pws.Range('A1').Font.Bold=$true
$pws.Columns.Item(1).ColumnWidth=60
$pws.Range('B:B').ColumnWidth=14
$pt.TableRange1.Columns.AutoFit() | Out-Null
$pws.Columns.Item(1).ColumnWidth=60
$wbK.Save()
L ("Pivot: {0}" -f $pt.TableRange1.Address($false,$false))

# ---------- 3. NHÓM HÓA ĐƠN ----------
$invs=@{}
foreach($ln in $lines){ if(-not $invs.ContainsKey($ln.inv)){ $invs[$ln.inv]=[pscustomobject]@{inv=$ln.inv;amt=0.0;wd=$ln.wd} }; $invs[$ln.inv].amt+=$ln.amt }
$reTG='NGÀY\s+(\d{1,2}/\d{1,2}/\d{4})\s*-\s*(.+?)\s*-\s*(TGHS|TG)\s*-'
foreach($iv in $invs.Values){
  $m=[regex]::Match($iv.wd,$reTG)
  if($m.Success){ $iv | Add-Member typ $m.Groups[3].Value; $iv | Add-Member name $m.Groups[2].Value.Trim(); $iv | Add-Member tdate ([datetime]::ParseExact($m.Groups[1].Value,'d/M/yyyy',$ci)) }
  else{ $iv | Add-Member typ 'LE'; $iv | Add-Member name ($iv.wd.Trim().Trim('"')); $iv | Add-Member tdate $null }
}
# ---------- 4. ĐƠN TRÌNH ----------
$trinh=@()  # typ,name,phone,addr,sp,weight,money,src
$tdates=$invs.Values | ? { $_.tdate } | % { $_.tdate } | Sort-Object -Unique
if(-not $tdates){ $tdates=@() }
foreach($td in $tdates){
  $tag=$td.ToString('dd.MM.yyyy')
  foreach($def in @(@('TG',"GV\DON HANG TRO GIA MT $tag.xlsx",3,9,44,42,48,28), @('TGHS',"HS\DON HANG TGHS $tag.xlsx",3,9,33,32,37,25))){
    $p=Join-Path $root $def[1]
    if(-not (Test-Path $p)){ L "Khong thay file: $p"; continue }
    $wb=$xl.Workbooks.Open($p,0,$true); $t=ReadSheet $wb 'HẰNG NGÀY'
    $hr=0; for($r=$t.r0;$r -lt $t.r0+$t.nr;$r++){ if((S (Cell $t $r 1)) -eq 'STT'){ $hr=$r; break } }
    for($r=$hr+1;$r -lt $t.r0+$t.nr;$r++){
      $stt=Cell $t $r 1; if($null -eq $stt -or -not ("$stt" -match '^\d+$')){ break }
      $trinh+=[pscustomobject]@{typ=$def[0];date=$td;name=(S (Cell $t $r $def[2]));phone=(Phone (Cell $t $r $def[3]));addr=(S (Cell $t $r $def[4]));sp=((Cell $t $r $def[5]) -as [string]).TrimEnd();wt=[double](Cell $t $r $def[6]);money=[double](Cell $t $r $def[7]);tinh=(S (Cell $t $r 6));used=$false}
    }
    $wb.Close($false)
  }
}
# đơn lẻ
$leTag= if($tdates.Count){ $tdates[0].ToString('dd.MM.yyyy') } else { '' }
$lef=Join-Path $root "Đơn lẻ\TRÌNH ĐƠN LẺ $leTag.xlsx"
if(Test-Path $lef){
  $wb=$xl.Workbooks.Open($lef,0,$true); $t=ReadSheet $wb 2
  for($r=$t.r0;$r -lt $t.r0+$t.nr;$r++){
    $nm=S (Cell $t $r 2); $mn=Cell $t $r 26; if($nm -eq '' -or $nm -match 'Tên KH' -or -not ($mn -is [double]) -or $mn -le 0){ continue }
    $full=S (Cell $t $r 1); $ph=''; $mm=[regex]::Match($full,'(?<!\d)0\d{9}(?!\d)'); if($mm.Success){ $ph=$mm.Value }
    $trinh+=[pscustomobject]@{typ='LE';date=$null;name=$nm;phone=$ph;addr=(S (Cell $t $r 3));sp=((Cell $t $r 27) -as [string]).TrimEnd();wt=[double](Cell $t $r 24);money=[double]$mn;tinh='';used=$false}
  }
  $wb.Close($false)
} else { L "Khong thay file don le: $lef" }
L ("Don trinh doc duoc: GV={0} HS={1} LE={2}" -f ($trinh|?{$_.typ -eq 'TG'}).Count,($trinh|?{$_.typ -eq 'TGHS'}).Count,($trinh|?{$_.typ -eq 'LE'}).Count)

# ---------- 5. GHÉP ----------
function Norm($s){ ($s -replace '\s+',' ').Trim().ToLower() }
$groups=$invs.Values | Group-Object { $_.typ + '|' + (Norm $_.name) }
$tem=@()
foreach($g in $groups){
  $typ=$g.Group[0].typ; $nm=Norm $g.Group[0].name
  $cands=@($trinh | ? { $_.typ -eq $typ -and ((Norm $_.name) -eq $nm -or ($typ -eq 'LE' -and ($nm.Contains((Norm $_.name)) -or (Norm $_.name).Contains($nm)))) -and (-not $_.used) })
  if($cands.Count -eq 0){ L "KHONG TIM THAY don trinh: [$typ] $($g.Group[0].name) - HD: $(($g.Group|%{$_.inv}) -join ',')"; continue }
  $ivs=@($g.Group | Sort-Object inv)
  $assign=@{}   # inv -> cand index
  if($cands.Count -eq 1){ foreach($iv in $ivs){ $assign[$iv.inv]=0 } }
  else{
    $n=$ivs.Count; $kk=$cands.Count; $total=[math]::Pow($kk,$n); $found=$false
    for($mask=0;$mask -lt $total -and -not $found;$mask++){
      $sum=@(0.0)*$kk; $x=$mask; $as=@()
      for($i=0;$i -lt $n;$i++){ $d=$x % $kk; $x=[math]::Floor($x/$kk); $sum[$d]+=$ivs[$i].amt; $as+=$d }
      $ok=$true; for($j=0;$j -lt $kk;$j++){ if([math]::Abs($sum[$j]-$cands[$j].money) -gt 1){$ok=$false} }
      if($ok){ $found=$true; for($i=0;$i -lt $n;$i++){ $assign[$ivs[$i].inv]=$as[$i] } }
    }
    if(-not $found){ L "CANH BAO: khong chia duoc hoa don theo so tien cho [$typ] $($g.Group[0].name) - gan het cho don dau"; foreach($iv in $ivs){ $assign[$iv.inv]=0 } }
  }
  for($j=0;$j -lt $cands.Count;$j++){
    $mine=@($ivs | ? { $assign[$_.inv] -eq $j })
    if($mine.Count -eq 0){ L "Don trinh khong co hoa don: [$typ] $($cands[$j].name)"; continue }
    $c=$cands[$j]; $c.used=$true
    $sumInv=($mine | Measure-Object amt -Sum).Sum
    if([math]::Abs($sumInv-$c.money) -gt 1){ L ("LECH TIEN: {0} HD {1}: bang ke {2:N0} vs don trinh {3:N0}" -f $c.name,(($mine|%{$_.inv}) -join '+'),$sumInv,$c.money) }
    $ids=($mine | %{$_.inv}) -join '+'
    $bill= switch($typ){ 'TG' {"$ids-TG-$($c.name)"} 'TGHS' { if($c.sp -match '(?i)PLUS'){"$ids-PLUS-$($c.name)"}else{"$ids-TGHS-$($c.name)"} } default {"$ids-$($c.name)"} }
    $hcm= ("$($c.addr) $($c.tinh)" -match '(?i)(h[oồ]\s*ch[ií]\s*minh|\bTP\.?\s*HCM\b|\bTPHCM\b|\bHCM\b|\bTp\s*HCM\b)')
    $tem+=[pscustomobject]@{first=[int]$mine[0].inv;bill=$bill;phone=$c.phone;addr=$c.addr;sp=$c.sp;wt=$c.wt;money=$c.money;hcm=$hcm}
  }
}
foreach($c in ($trinh | ? { -not $_.used })){ L "Don trinh chua co hoa don trong bang ke: [$($c.typ)] $($c.name)" }
$tem=@($tem | Sort-Object first)
L "Tem tao: $($tem.Count) don; tong COD = $(($tem | Measure-Object money -Sum).Sum)"
L "Tong bang ke = $(($invs.Values | Measure-Object amt -Sum).Sum)"

# ---------- 6. GHI FILE TEM ----------
$outPath=Join-Path $temDir $outName
Copy-Item $baseTemPath $outPath -Force
$wbT=$xl.Workbooks.Open($outPath); $ws=$wbT.Worksheets.Item('Danh sách')
$ws.Range('A8:AF200').ClearContents() | Out-Null
$n=$tem.Count
if($n -gt 1){ $ws.Range('A8:V8').Copy() | Out-Null; $ws.Range("A9:V$(7+$n)").PasteSpecial(-4122) | Out-Null }
$ws.Range("D8:D$(7+$n)").NumberFormat='@'
for($i=0;$i -lt $n;$i++){
  $r=8+$i; $tt=$tem[$i]
  $ws.Cells.Item($r,1).Value2=[double]($i+1)
  $ws.Cells.Item($r,3).Value2=[string]$tt.bill
  $ws.Cells.Item($r,4).Value2=[string]$tt.phone
  $ws.Cells.Item($r,5).Value2=[string]$tt.addr
  $ws.Cells.Item($r,6).Value2=[string]$tt.sp
  $ws.Cells.Item($r,7).Value2=1
  $ws.Cells.Item($r,8).Value2=[double]$tt.wt
  $ws.Cells.Item($r,9).Value2=[double]$tt.money
  $ws.Cells.Item($r,10).Value2=[double]$tt.money
  $ws.Cells.Item($r,11).Value2='Bưu kiện'
  if($tt.hcm){ $ws.Cells.Item($r,13).Value2='PHS - Nội tỉnh tiết kiệm thỏa thuận' }
  else{ $ws.Cells.Item($r,13).Value2='LCOD - TMĐT tiết kiệm thỏa thuận'; $ws.Cells.Item($r,14).Value2='GDK' }
  $ws.Cells.Item($r,19).Value2='Người nhận trả'
  $ws.Cells.Item($r,20).Value2='Gọi điện thoại trước khi giao'
  $ws.Cells.Item($r,22).Value2='Giờ hành chính'
}
$ws.Range("A8:V$(7+$n)").Rows.AutoFit() | Out-Null
$wbT.Save(); $wbT.Close($false)
L "Da ghi: $outPath"
$wbK.Close($false)
} finally { $xl.Quit() }
if($report){ $log | Set-Content $report -Encoding utf8 }
