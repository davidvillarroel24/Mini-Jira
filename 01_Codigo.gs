function doGet() {

  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('Mini Jira QA')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}

function getView(name) {

  return HtmlService
    .createHtmlOutputFromFile(name)
    .getContent();
}