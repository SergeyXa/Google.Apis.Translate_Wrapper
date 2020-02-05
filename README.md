# Google.Apis.Translate Wrapper

## What's this?
This is a tiny .NET library, which employs [Google Cloud
Translation](https://cloud.google.com/translate/docs/) for translating a text from one
language to another. 

The project depends on the [Google.Apis](https://www.nuget.org/packages/Google.Apis),
mainly on the
[Google.Apis.Translate.v3beta1](https://www.nuget.org/packages/Google.Apis.Translate.v3beta1),
and, in fact, is just a minimalistic wrapper around these, making it simple to use the
APIs for a text translation from .NET or any environment supporting COM.

The translator can be used from VBA (e.g. in Word, or Excel). The project was initially 
developed to translate some Word documents. The source of the macro demonstrating usage 
of the translator in Microsoft Word is included.

## Building And Installation

### Obtaining the `key.json` file from Google

To use the translation api you should register the project in the [Cloud
console](https://console.cloud.google.com/).

- [Open console](https://console.cloud.google.com/)
- Create a project
- [Enable Cloud Translation API](https://console.cloud.google.com/apis/library) for the
  project
- [Create a service account](https://console.cloud.google.com/apis/credentials) for the
  project
- Open the [service accounts
  page](https://console.cloud.google.com/iam-admin/serviceaccounts) and choose to
  generate a key in actions of the account you have created.
- Place the downloaded file to the project tree, near the `Translator.cs`, as `key.json`.

The key.json file bounds the usage of the project to your own service account. Notice
that after some limit, usage of the Cloud Translation API is paid. So be careful when
publishing your own key.json. All the risks will be on you.

The compilation process simply copies the key.json file to the output folder. So you can
replace the key with another one even when the library is already compiled.

### Compiling And Installing

Fetch the source, open the solution on VS and compile. I developed the project on VS
2017, though most other versions should be capable to compile the code as well.

Notice that if you are going to use the library via COM, you have to run the Visual
Studio as an administrator. Otherwise the VS won't register the library after it's
compiled.

### Is it free to use the Google Cloud Translation?

At the moment of this writing, Google allows to use the API for free to translate up to
500 000 characters monthly. It should be very enough for personal use. For the characters
above, you will have to pay. See [Pricing](https://cloud.google.com/translate/pricing)
page for details.
