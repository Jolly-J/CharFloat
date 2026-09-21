function Get-OfficeApp($component) {
  $class = @{excel='Excel.Application';word='Word.Application';ppt='PowerPoint.Application'}[$component]
  if (!$class) { throw 'Unknown Office component' }
  try { return [Runtime.InteropServices.Marshal]::GetActiveObject($class) }
  catch { throw "Cannot attach to $class. Open the desktop app and a document in the same user session. $($_.Exception.Message)" }
}
function Get-Target($app, $component, $name) {
  $collection=$null
  switch ($component) { 'excel' {$collection=$app.Workbooks} 'word' {$collection=$app.Documents} 'ppt' {$collection=$app.Presentations} default {throw 'Unknown component'} }
  if ($name) {
    $found=@()
    for($i=1;$i -le $collection.Count;$i++) {$item=$collection.Item($i);if($item.Name -ceq $name -or $item.FullName -ceq $name){$found+=,$item}}
    if ($found.Count -ne 1) { throw "Target must match exactly one open document: $name" }
    return $found[0]
  }
  if ($collection.Count -ne 1) { throw 'Specify a document name when zero or multiple documents are open' }
  return $collection.Item(1)
}
