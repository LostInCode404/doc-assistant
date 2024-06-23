// Constants
OPENAI_COMPLETIONS_API_URL = 'https://api.openai.com/v1/chat/completions';

// Supported features
const features = {
  GENERATE_IDEAS: 'ideas',
  GENERATE_ESSAY: 'essay',
  SUMMARIZE_TEXT: 'summarize',
  TWITTER_POST: 'twitter-post',
  LINKEDIN_POST: 'linkedin-post',
  INSTAGRAM_POST: 'instagram-post'
}
const featuresPromptMap = {
  [features.GENERATE_IDEAS]: 'Generate 10 ideas on:\n',
  [features.GENERATE_ESSAY]: 'Generate an essay containing upto 400 words on:\n',
  [features.SUMMARIZE_TEXT]: 'Summarize the following text:\n',
  [features.TWITTER_POST]: 'Generate 5 twitter post ideas for:\n',
  [features.LINKEDIN_POST]: 'Generate 5 linkedin post ideas for:\n',
  [features.INSTAGRAM_POST]: 'Generate 5 instagram post ideas for:\n',
}

// Get OpenAI params on script load
const openAiProps = getOpenAiProps();

// Add menu item and set handler function name for it
function onOpen() {
  DocumentApp
    .getUi()
    .createMenu('DocAssistant')
    .addItem('Generate Ideas', "generateIdeas")
    .addItem('Generate Essay', "generateEssay")
    .addItem('Summarize Text', "summarizeText")
    .addSubMenu(
      DocumentApp.getUi().createMenu('Generate Social Media Post')
        .addItem('Twitter', 'generateTwitterPosts')
        .addItem('Linkedin', 'generateLinkedinPosts')
        .addItem('Instagram', 'generateInstagramPosts')
      )
    .addToUi();
}

// Helper functions to generate content
function generateIdeas() {generateContent(features.GENERATE_IDEAS);}
function generateEssay() {generateContent(features.GENERATE_ESSAY);}
function summarizeText() {generateContent(features.SUMMARIZE_TEXT);}
function generateTwitterPosts() {generateContent(features.TWITTER_POST);}
function generateLinkedinPosts() {generateContent(features.LINKEDIN_POST);}
function generateInstagramPosts() {generateContent(features.INSTAGRAM_POST);}

// Generate text using the selected text and add it to the document
function generateContent(feature) {

  // Get prompt data from the document
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  let selectedText = '';
  if(selection) {
    selectedText = selection.getRangeElements()[0].getElement().asText().getText();
  } else {
    throw new Error('No text selected')
  }

  // Generate prompt based on requested feature
  const prompt = featuresPromptMap[feature] + selectedText;

  // Make ChatGPT API call
  const temperature = 0.2;
  const maxTokens = 1024;
  const requestBody = {
    model: openAiProps.model,
    messages: [{role: "user", content: prompt}],
    temperature,
    max_tokens: maxTokens
  }
  const requestOptions = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: 'Bearer ' + openAiProps.apiKey
    },
    payload: JSON.stringify(requestBody)
  }
  const response = UrlFetchApp.fetch(OPENAI_COMPLETIONS_API_URL, requestOptions);
  const responseBody = JSON.parse(response.getContentText());
  const generatedText = responseBody.choices[0].message.content;
  
  // Append result to document body
  const docBody = doc.getBody();
  docBody.appendParagraph(`\n${generatedText}`);
}

// Function to get OpenAI API params
function getOpenAiProps() {
  const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  const OPENAI_MODEL = PropertiesService.getScriptProperties().getProperty('OPENAI_MODEL') || 'gpt-3.5-turbo';
  return {
    apiKey: OPENAI_API_KEY,
    model: OPENAI_MODEL
  }
}
