function add(first, second){
  return first + second;
}

function increment(incrementBy, callback) {
  var result = 0;
  var timer = setInterval(function() {
    result += incrementBy;
    callback.setResult(result);
  }, 1000);

  callback.onCanceled = function() {
    clearInterval(timer);
  };
}

//JSONメタデータで定義した「ADD」に関数「add」が対応するように指定
CustomFunctionMappings.ADD = add;
//JSONメタデータで定義した「INCREMENT」に関数「increment」が対応するように指定
CustomFunctionMappings.INCREMENT = increment;