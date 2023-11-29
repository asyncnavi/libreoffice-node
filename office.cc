#include "office.h"
#include <napi.h>
#include <stdlib.h>
#include <iostream>
#include "LibreOfficeKit/LibreOfficeKit.h"
#include "LibreOfficeKit/LibreOfficeKitInit.h"

Office::Office(const Napi::CallbackInfo& info)
    : Napi::ObjectWrap<Office>(info) {}

void Office::Load(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (info.Length() < 1) {
    Napi::TypeError::New(env, "Load: Wrong number of arguments")
        .ThrowAsJavaScriptException();
    return;
  }

  if (!info[0].IsString()) {
    Napi::TypeError::New(env, "Load: Path must be string")
        .ThrowAsJavaScriptException();
    return;
  }
  std::string path = info[0].As<Napi::String>().Utf8Value();

  LibreOfficeKit* office = lok_init(path.c_str());
  if (office == NULL) {
    Napi::TypeError::New(env, "Load: Failed to load libreofficekit")
        .ThrowAsJavaScriptException();
    return;
  }
  this->SetOffice(office);
}

void Office::Destroy(const Napi::CallbackInfo& info) {
  if (this->office == NULL) {
    return;
  }
  this->office->pClass->destroy(this->office);
}

void Office::LoadDocument(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (this->office == NULL) {
    Napi::TypeError::New(env, "LoadDocument: Libreoffice library is not loaded")
        .ThrowAsJavaScriptException();
    return;
  }

  if (info.Length() < 1) {
    Napi::TypeError::New(env, "LoadDocument: Wrong number of arguments")
        .ThrowAsJavaScriptException();
    return;
  }

  if (!info[0].IsString()) {
    Napi::TypeError::New(env, "LoadDocument: Path must be string")
        .ThrowAsJavaScriptException();
    return;
  }
  std::string path = info[0].As<Napi::String>().Utf8Value();
  LibreOfficeKitDocument* doc =
      this->office->pClass->documentLoad(this->office, path.c_str());
  if (doc == NULL) {
    char* err_msg = this->office->pClass->getError(this->office);
    if (err_msg != NULL) {
      Napi::TypeError::New(env, strcat("LoadDocument:", err_msg))
          .ThrowAsJavaScriptException();
      return;
    }
    Napi::TypeError::New(env, "LoadDocument: Unable to load document")
        .ThrowAsJavaScriptException();
    return;
  }

  doc->pClass->destroy(doc);
  std::cout << "done" << std::endl;
}

Napi::Object Office::Init(Napi::Env env, Napi::Object exports) {
  Napi::Function func =
      DefineClass(env,
                  "Office",
                  {InstanceMethod("load", &Office::Load),
                   InstanceMethod("load_document", &Office::LoadDocument),
                   InstanceMethod("destroy", &Office::Destroy)});

  Napi::FunctionReference* constructor = new Napi::FunctionReference();
  *constructor = Napi::Persistent(func);
  env.SetInstanceData(constructor);

  exports.Set("Office", func);
  return exports;
}

Document::Document(const Napi::CallbackInfo& info,
                   LibreOfficeKitDocument* document)
    : Napi::ObjectWrap<Document>(info) {
  this->document = document;
}

Napi::Object Document::Init(Napi::Env env, Napi::Object exports) {
  Napi::Function func =

      DefineClass(env, "Document", {});

  Napi::FunctionReference* constructor = new Napi::FunctionReference();
  *constructor = Napi::Persistent(func);
  env.SetInstanceData(constructor);

  exports.Set("Document", func);
  return exports;
}

Napi::Object Init(Napi::Env env, Napi::Object exports) {
  Document::Init(env, exports);
  return Office::Init(env, exports);
}

NODE_API_MODULE(addon, Init)
