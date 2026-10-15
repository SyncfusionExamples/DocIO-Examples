using Azure.AI.OpenAI;
using OpenAI.Chat;
using Syncfusion.DocIO.DLS;
using System.ClientModel;

namespace WordTranslatorWithAzureOpenAI
{
    internal class Program
    {
        // Replace with your Azure OpenAI API key.
        static string? azureOpenAIApiKey = "Replace your Azure OpenAI key";

        // Replace with your Azure OpenAI deployment name.
        static string? azureOpenAIModel = "Replace your deployment name";

        static async Task Main(string[] args)
        {
            // Start the Word document translation process.
            await ExecuteTranslation();
        }

        /// <summary>
        /// Execute translation of Word document.
        /// </summary>
        private async static Task ExecuteTranslation()
        {
            // Display application title.
            Console.WriteLine("AI Powered Word Translator");

            // Prompt the user to enter the full path of the Word document.
            Console.WriteLine("Enter full Word file path (e.g., C:\\Data\\Input.docx):");

            // Read the Word document path entered by the user.
            string? wordFilePath = Console.ReadLine()?.Trim().Trim('"');

            // Prompt the user to enter the target language.
            Console.WriteLine("Enter the language name (e.g., Chinese, Japanese)");

            // Read the target language entered by the user.
            string? language = Console.ReadLine()?.Trim().Trim('"');

            // Validate the specified Word document path.
            if (string.IsNullOrWhiteSpace(wordFilePath) || !File.Exists(wordFilePath))
            {
                Console.WriteLine("Invalid path. Exiting.");
                return;
            }

            // Validate the target language.
            if (string.IsNullOrWhiteSpace(language))
            {
                Console.WriteLine("Invalid language. Exiting.");
                return;
            }

            // Verify that the Azure OpenAI API key is configured.
            if (string.IsNullOrWhiteSpace(azureOpenAIApiKey))
            {
                Console.WriteLine("AZURE_OPENAI_API_KEY not set. Exiting.");
                return;
            }

            try
            {
                // Generate and save the translated Word document.
                await TranslateWordContent(wordFilePath, language);
            }
            catch (Exception ex)
            {
                // Handle any errors encountered during translation.
                Console.WriteLine($"Failed to read Word document: {ex.StackTrace}");
                return;
            }
        }

        /// <summary>
        /// Translate Word content using Azure OpenAI and Syncfusion DocIO.
        /// </summary>
        /// <param name="wordFilePath">Path to the Word file.</param>
        /// <param name="language">Target language for translation.</param>
        private static async Task TranslateWordContent(string wordFilePath, string language)
        {
            // Open the existing Word document.
            WordDocument wordDocument = new WordDocument(wordFilePath);

            // Define a system prompt that instructs Azure OpenAI how to perform the translation.
            // The prompt includes rules to preserve formatting, placeholders, and document structure.
            string systemPrompt = @"You are a professional translator integrated into an DocIO automation tool.
                                    Your job is to translate text from word document into the" + language + @" language
                                    Rules:
                                    - Preserve original structure as much as possible.
                                    - Prefer literal translation over paraphrasing.
                                    - Return ONLY the translated text, without quotes, labels, or explanations.
                                    - Preserve placeholders (e.g., {0}, {name}) and keep numbers, currency, and dates intact.
                                    - Do not change the meaning, tone, or formatting unnecessarily.
                                    - Do not add extra commentary or code fences.
                                    - If the text is already in the target language, return it unchanged.
                                    - Be concise and accurate.";

            // Iterate through each section in the Word document.
            foreach (WSection section in wordDocument.Sections)
            {
                // Access the section body, which contains the main document content.
                WTextBody sectionBody = section.Body;

                // Translate the content in the section body.
                await TranslateTextBody(sectionBody, systemPrompt);

                // Access the headers and footers of the current section.
                WHeadersFooters headersFooters = section.HeadersFooters;

                // Translate the content in the odd-page header and footer.
                await TranslateTextBody(headersFooters.OddHeader, systemPrompt);
                await TranslateTextBody(headersFooters.OddFooter, systemPrompt);
            }

            // Save the Translate Word document.
            wordDocument.Save(wordFilePath.Replace(".docx", "_DocIOTranslate.docx"));

            // Close the Word document and release resources.
            wordDocument.Close();
        }

        /// <summary>
        /// Translates the text content of a text body using Azure OpenAI and updates it.
        /// </summary>
        /// <param name="textBody">The textBody to translate.</param>
        /// <param name="systemPrompt">System prompt to guide translation behavior.</param>
        private static async Task TranslateTextBody(WTextBody textBody, string systemPrompt)
        {
            // Iterate through each content element in the text body.
            foreach (Entity entity in textBody.ChildEntities)
            {
                // Translate paragraph content.
                if (entity is WParagraph paragraph)
                {
                    await TranslateParagraphs(paragraph, systemPrompt);
                }

                // Translate table content.
                else if (entity is WTable table)
                {
                    await TranslateTable(table, systemPrompt);
                }

                // Translate content within content controls recursively.
                else if (entity is BlockContentControl blockContentControl)
                {
                    await TranslateTextBody(blockContentControl.TextBody, systemPrompt);
                }
            }
        }

        /// <summary>
        /// Translates all text content within a table using OpenAI.
        /// </summary>
        /// <param name="table">The table to translate.</param>
        /// <param name="systemPrompt">System prompt to guide translation behavior.</param>
        private static async Task TranslateTable(WTable table, string systemPrompt)
        {

            // Iterate through each row in the table.
            foreach (WTableRow row in table.Rows)
            {
                // Iterate through each cell in the current row.
                foreach (WTableCell cell in row.Cells)
                {
                    // Iterate through each content element within the cell.
                    foreach (Entity entity in cell.ChildEntities)
                    {
                        // Translate paragraph content.
                        if (entity is WParagraph paragraph)
                        {
                            await TranslateParagraphs(paragraph, systemPrompt);
                        }

                        // Translate paragraph content.
                        else if (entity is WTable nestedTable)
                        {
                            await TranslateTable(nestedTable, systemPrompt);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Translates the text content of a paragraph using Azure OpenAI and updates it.
        /// </summary>
        /// <param name="paragraph">The paragraph to translate.</param>
        /// <param name="systemPrompt">System prompt to guide translation behavior.</param>
        private static async Task TranslateParagraphs(WParagraph paragraph, string systemPrompt)
        {
            // Retrieve the paragraph text.
            string originalText = paragraph.Text;
            if (string.IsNullOrEmpty(originalText)) return;
            try
            {
                string translatedText = originalText;
                // Call Azure OpenAI to translate the text and Store the translated result in the dictionary for reuse.
                translatedText = await AskAzureOpenAIAsync(azureOpenAIApiKey, azureOpenAIModel, systemPrompt, originalText);
                // Replace the original text with the translated version.
                paragraph.Text = translatedText;
            }
            catch (Exception ex)
            {
                // Log any errors that occur during translation.
                Console.WriteLine($"OpenAI error: {ex.Message}");
            }
        }

        /// <summary>
        /// Sends a chat completion request to Azure OpenAI and returns the response.
        /// </summary>
        /// <param name="apiKey">Azure OpenAI API key.</param>
        /// <param name="azureOpenAImodel">Model name.</param>
        /// <param name="systemPrompt">System prompt.</param>
        /// <param name="userContent">User content.</param>
        /// <returns>AI-generated response as a string.</returns>
        private static async Task<string> AskAzureOpenAIAsync(string apiKey, string azureOpenAImodel, string systemPrompt, string userContent)
        {
            // Create the Azure OpenAI client using the endpoint and API key
            AzureOpenAIClient azureClient = new(
                new Uri("https://your-resource-name.openai.azure.com/"),
                new ApiKeyCredential(apiKey));

            // Create chat client for the specified mode
            ChatClient chatClient = azureClient.GetChatClient(azureOpenAIModel);

            // Create chat client for the specified mode
            ClientResult<ChatCompletion> chatResult = await chatClient.CompleteChatAsync(
                new SystemChatMessage(systemPrompt),
                new UserChatMessage(userContent));

            // Extract the generated content from the response
            string response = chatResult.Value.Content[0].Text ?? string.Empty;
            return response;
        }
    }
}
