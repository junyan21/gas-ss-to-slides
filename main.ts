function createPresentation() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var presentation = SlidesApp.create('GASで作ったスライド');
    var slides = presentation.getSlides();

    // 1枚目のスライド
    var title = sheet.getName();
    var slide = slides[0] || presentation.appendSlide();
    var textBox = slide.getShapes()[0] || slide.insertShape(SlidesApp.ShapeType.TEXT_BOX);
    textBox.getFill().setTransparent();
    textBox.getBorder().setTransparent();
    textBox.setLeft(50);
    textBox.setTop(50);
    textBox.setWidth(400);
    textBox.setHeight(50);
    var textRange = textBox.getText();
    textRange.setText(title);

    // 2枚目以降のスライド
    var range = sheet.getRange("A1:A");
    var values = range.getValues();
    for (var i = 0; i < values.length; i++) {
        var slide = slides[i + 1] || presentation.appendSlide();
        slide.getLines()[0].setDashStyle(SlidesApp.DashStyle.SOLID);
        // スライドの枠線描画
        slide.getShapes()[0].getBorder().setWeight(1).setDashStyle(SlidesApp.DashStyle.SOLID);

        // タイトル用のテキストボックスを設定する
        var textBox = slide.getShapes()[0] || slide.insertShape(SlidesApp.ShapeType.TEXT_BOX);
        setupSlideTitleTextBox(textBox, values[i][0]);
    }
}

// スライドタイトル用のテキストボックスを設定する
function setupSlideTitleTextBox(textBox: GoogleAppsScript.Slides.Shape, text: string) {
    textBox.getFill().setTransparent();
    textBox.getBorder().setTransparent();
    textBox.setTop(30);
    textBox.setWidth(400);
    textBox.setHeight(30);
    // テキストBox自体を左右中央揃え
    textBox.alignOnPage(SlidesApp.AlignmentPosition.HORIZONTAL_CENTER)
    var textRange = textBox.getText();
    textRange.setText(text);
    // テキストボックス内のテキストを左右中央寄せ
    textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    // テキストのスタイル設定
    let textStyle = textRange.getTextStyle();
    textStyle.setFontFamily('Arial');
    textStyle.setBold(true);
    textStyle.setFontSize(20);
    textStyle.setForegroundColor('#2229f3');
}