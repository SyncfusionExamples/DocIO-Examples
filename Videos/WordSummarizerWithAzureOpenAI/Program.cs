using Azure.AI.OpenAI;
using OpenAI.Chat;
using Syncfusion.DocIO.DLS;
using System.ClientModel;

namespace WordSummarizerWithAzureOpenAI
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            // Replace with your Azure OpenAI API key
            string? azureOpenAIApiKey = "Replace your Azure OpenAI key";

            // Start the Word document summarization process
            await ExcuteSummarization(azureOpenAIApiKey);
        }

        /// <summary>
        /// Execute summarization of Word document.
        /// </summary>
        private async static Task ExcuteSummarization(string azureOpenAIApiKey)
        {
            // Display application title
            Console.WriteLine("AI Powered Word Summarizer");

            // Prompt user for the input Word document path
            Console.WriteLine("Enter full Word file path (e.g., C:\\Data\\Input.docx):");
            string? wordFilePath = Console.ReadLine()?.Trim().Trim('"');

            // Prompt user for the desired summary length
            Console.WriteLine("Please enter the number of sentences you would like the summary to be (e.g., 3, 5):");
            string? sentencesCount = Console.ReadLine()?.Trim().Trim('"');

            // Validate the Word document path
            if (string.IsNullOrWhiteSpace(wordFilePath) || !File.Exists(wordFilePath))
            {
                Console.WriteLine("Invalid path. Exiting.");
                return;
            }

            // Validate the sentence count input
            if (string.IsNullOrWhiteSpace(sentencesCount) || !int.TryParse(sentencesCount, out int result))
            {
                Console.WriteLine("Invalid Count. Exiting.");
                return;
            }

            // Ensure the Azure OpenAI API key is available
            if (string.IsNullOrWhiteSpace(azureOpenAIApiKey))
            {
                Console.WriteLine("AZURE_OPENAI_API_KEY not set. Exiting.");
                return;
            }

            try
            {
                // Generate and save the summarized Word document
                await SummarizeWordContent(azureOpenAIApiKey, wordFilePath, sentencesCount);
            }
            catch (Exception ex)
            {
                // Handle summarization errors
                Console.WriteLine($"Failed to summarize Word document: {ex.StackTrace}");
                return;
            }
        }

        /// <summary>
        /// Reads the content of a Word document, generates a summary using Azure OpenAI,
        /// and saves the summarized content as a new Word document.
        /// </summary>
        private static async Task SummarizeWordContent(string azureOpenAIApiKey, string wordFilePath, string sentencesCount)
        {
            // Load the source Word document
            WordDocument wordDocument = new WordDocument(wordFilePath);

            // Create a prompt instructing the AI to summarize the content
            string systemPrompt = @"You are a professional document summarizer integrated into an DocIO automation tool.
                                    Your job is to summarize the word document content into the"" + sentencesCount + "" sentences";
            
            // Extract all text from the document
            string originalText = wordDocument.GetText();

            // Close the source document after reading
            wordDocument.Close();

            // Send document content to Azure OpenAI and get the summary
            string summarizedText = await AskAzureOpenAIAsync(azureOpenAIApiKey, systemPrompt, originalText);

            // Create a new Word document to store the summary
            WordDocument summarizedDocument = new WordDocument();
            summarizedDocument.EnsureMinimal();

            // Add the summarized text to the document
            summarizedDocument.LastParagraph.AppendText(summarizedText);

            // Save the summarized document with a new file name
            summarizedDocument.Save(wordFilePath.Replace(".docx", "_DocIOsummarized.docx"));

            // Close the summarized document
            summarizedDocument.Close();
        }

        /// <summary>
        /// Sends a chat completion request to OpenAI and returns the response.
        /// </summary>
        /// <param name="apiKey">Azure OpenAI API key.</param>
        /// <param name="model">Model name.</param>
        /// <param name="systemPrompt">System prompt.</param>
        /// <param name="userContent">User content.</param>
        /// <returns>AI-generated response as a string.</returns>
        private static async Task<string> AskAzureOpenAIAsync(string apiKey, string systemPrompt, string userContent)
        {
            // Create the Azure OpenAI client using the endpoint and API key
            AzureOpenAIClient azureClient = new(
                new Uri("https://your-resource-name.openai.azure.com/"),
                new ApiKeyCredential(apiKey)
                );

            // Create chat client for the specified mode
            ChatClient chatClient = azureClient.GetChatClient("your-model-name");
            
            // Send the system prompt and document content to the model
            ClientResult<ChatCompletion> chatResult = await chatClient.CompleteChatAsync(
                new SystemChatMessage(systemPrompt),
                new UserChatMessage(userContent));
            
            // Extract the generated summary from the response
            string response = chatResult.Value.Content[0].Text ?? string.Empty;
            return response; 
        }
    }
}
