URL = "https://docs.google.com/forms/d/1tI-zUCSdfDIVYSmHiyx7S8zxlb0SlWq4Vki2FtYGsKY/edit"

function countGridResponses() {
  // フォームIDを設定

  var form = FormApp.openByUrl(URL);
  var formResponses = form.getResponses();

  // グリッド質問のタイトルを設定
  var questionTitle = '希望日2';

  // 回答データを格納するオブジェクト
  var responseData = {};

  // 全てのフォームの回答をループ
  formResponses.forEach(function(formResponse) {
    var itemResponses = formResponse.getItemResponses();
    
    // 各質問に対する回答をループ
    itemResponses.forEach(function(itemResponse) {
      if (itemResponse.getItem().getTitle() === questionTitle) {
        // グリッド質問の回答を取得
        var answers = itemResponse.getResponse();
        
        // グリッドの各行に対する回答を集計
        answers.forEach(function(answer, index) {
          var rowTitle = itemResponse.getItem().asGridItem().getRows()[index];
          if (!responseData[rowTitle]) {
            responseData[rowTitle] = 0; // 初期化
          }
          if (answer !== null) { // null以外の回答をカウント
            responseData[rowTitle]++;
          }
        });
      }
    });
  });

  // 集計結果をログに出力
  console.log(responseData);
}


function disableOptions() {
  var formId = 'YOUR_FORM_ID';
  var form = FormApp.openById(formId);
  var items = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE);

  items.forEach(function(item) {
    if (item.getTitle() === "質問のタイトル") {
      var mcItem = item.asMultipleChoiceItem();
      var choices = mcItem.getChoices();
      var newChoices = choices.map(function(choice) {
        if (choice.getValue() === "選択肢A") {
          // 選択肢に " - 利用不可" を追加
          return choice.getValue() + " - 利用不可";
        }
        return choice;
      });

      // 更新した選択肢で質問を更新
      mcItem.setChoices(newChoices);
    }
  });
}
