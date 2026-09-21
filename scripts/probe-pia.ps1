$asm = [System.Reflection.Assembly]::LoadFrom('C:\Users\rossj\.nuget\packages\exceldna.interop\16.0.0\lib\net6.0-windows7.0\Microsoft.Office.Interop.Excel.dll');
$ct = $asm.GetType('Microsoft.Office.Interop.Excel.Comment');
foreach ($m in $ct.GetMembers()) {
  if ($m.Name -in @('Text','Author','Visible','Delete','Shape')) {
    if ($m -is [System.Reflection.PropertyInfo]) {
      $p = [System.Reflection.PropertyInfo]$m;
      Write-Host "P  $($m.Name) :: $($p.PropertyType.FullName) get_virtual=$($p.GetMethod?.IsVirtual) set_virtual=$($null -ne $p.SetMethod -and $p.SetMethod.IsVirtual)";
    } elseif ($m -is [System.Reflection.MethodInfo]) {
      $mi = [System.Reflection.MethodInfo]$m;
      $ps = ($mi.GetParameters() | ForEach-Object { "$($_.Name):$($_.ParameterType.Name)" }) -join ';';
      Write-Host "M  $($m.Name) :: $($mi.ReturnType.FullName) virtual=$($mi.IsVirtual) static=$($mi.IsStatic) params=[$ps]";
    }
  }
}
Write-Host '--- Range.AddComment return + params (full) ---';
$rt = $asm.GetType('Microsoft.Office.Interop.Excel.Range');
$rt.GetMethod('AddComment').GetParameters() | ForEach-Object { Write-Host "$($_.Name): $($_.ParameterType.FullName) optional=$($_.IsOptional)" };
Write-Host '--- Range.Comment property ---';
$rt.GetProperty('Comment') | Select-Object Name, @{N='Type';E={$_.PropertyType.FullName}}, @{N='get_virtual';E={$_.GetMethod.IsVirtual}};
