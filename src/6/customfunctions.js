function add99(first, second) {
  return first + second + 99;
}
//JSONメタデータで定義した「ADD99」に関数「add99」が対応するように指定
CustomFunctions.associate("ADD99", add99);