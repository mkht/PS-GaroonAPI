function Get-RandomName () {
    # なんとなく市区町村ぽい名前をランダム生成する関数
    $MojiList1 = '上中下東西南北京浜丘国寺神和平右左日月火水木金土新古大小海池山河湖黒白一二三'
    $MojiList2 = '市区町村'
    $MojiA = Get-Random -InputObject $MojiList1.ToCharArray() -Count (Get-Random (2, 3))
    $MojiB = Get-Random -InputObject $MojiList2.ToCharArray() -Count 1
    -join ($MojiA + $MojiB)
}

Export-ModuleMember -Function Get-RandomName
