using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Translate.v3beta1;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace GoogleApisTranslateWrapper
{
    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    public class Translator
    {
        TranslateService _translateService;
        string _projectId;

        public Translator()
        {
            var keyJsonFilepath = Path.Combine(
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                    @"key.json");

            try
            {
                var credential = GoogleCredential.FromFile(keyJsonFilepath);

                if (!(credential.UnderlyingCredential is ServiceAccountCredential))
                    throw new Exception("key.json should define a ServiceAccountCredential");

                _projectId = ((ServiceAccountCredential)credential.UnderlyingCredential).ProjectId;

                if (credential.IsCreateScopedRequired)
                    credential = credential.CreateScoped(TranslateService.Scope.CloudTranslation);

                _translateService = new TranslateService(new BaseClientService.Initializer
                {
                    ApplicationName = nameof(GoogleApisTranslateWrapper),
                    HttpClientInitializer = credential
                });
            }
            catch (Exception ex)
            {

            }
        }

        public string Translate(string text, string sourceLang, string targetLang)
        {
            var translateTextRequest = _translateService.Projects.TranslateText(
                new Google.Apis.Translate.v3beta1.Data.TranslateTextRequest()
                {
                    Contents = new List<string>() { text },
                    SourceLanguageCode = sourceLang,
                    TargetLanguageCode = targetLang,
                }, 
                $"projects/{_projectId}");

            var translateTextResponse = translateTextRequest.Execute();

            return translateTextResponse.Translations[0].TranslatedText;
        }
    }
}
