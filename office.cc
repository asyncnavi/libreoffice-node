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
    Napi::TypeError::New(env, "wrong number of arguments")
        .ThrowAsJavaScriptException();
    return;
  }

  if (!info[0].IsString()) {
    Napi::TypeError::New(env, "path must be string")
        .ThrowAsJavaScriptException();
    return;
  }
  std::string path = info[0].As<Napi::String>().Utf8Value();

  LibreOfficeKit* office = lok_init(path.c_str());
  if (office == NULL) {
    Napi::TypeError::New(env, "failed to load libreofficekit")
        .ThrowAsJavaScriptException();
    return;
  }
  this->SetOffice(office);
}

void Office::Close(const Napi::CallbackInfo& info) {
  if (this->office == NULL) {
    return;
  }
  this->office->pClass->destroy(this->office);
}

void Office::Convert(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (this->office == NULL) {
    Napi::TypeError::New(env, "libreoffice library is not loaded")
        .ThrowAsJavaScriptException();
    return;
  }

  if (info.Length() < 3) {
    Napi::TypeError::New(env, "wrong number of arguments")
        .ThrowAsJavaScriptException();
    return;
  }
  if (!info[0].IsString()) {
    Napi::TypeError::New(env, "path must be string")
        .ThrowAsJavaScriptException();
    return;
  }
  if (!info[1].IsString()) {
    Napi::TypeError::New(env, "format must be string")
        .ThrowAsJavaScriptException();
    return;
  }
  if (!info[2].IsString()) {
    Napi::TypeError::New(env, "output path must be string")
        .ThrowAsJavaScriptException();
    return;
  }
  std::string path = info[0].As<Napi::String>().Utf8Value();
  std::string format = info[1].As<Napi::String>().Utf8Value();
  std::string output_path = info[2].As<Napi::String>().Utf8Value();

  LibreOfficeKitDocument* doc =
      this->office->pClass->documentLoad(this->office, path.c_str());
  if (doc == NULL) {
    char* err_msg = this->office->pClass->getError(this->office);
    doc->pClass->destroy(doc);
    if (err_msg != NULL) {
      Napi::TypeError::New(env, strcat("Convert:", err_msg))
          .ThrowAsJavaScriptException();
      return;
    }
    Napi::TypeError::New(env, "unable to save document")
        .ThrowAsJavaScriptException();
    return;
  }
  doc->pClass->destroy(doc);
}

Napi::Object Office::Init(Napi::Env env, Napi::Object exports) {
  Napi::Function func =
      DefineClass(env,
                  "Office",
                  {InstanceMethod("load", &Office::Load),
                   InstanceMethod("convert", &Office::Convert),
                   InstanceMethod("close", &Office::Close)});

  Napi::FunctionReference* constructor = new Napi::FunctionReference();
  *constructor = Napi::Persistent(func);
  env.SetInstanceData(constructor);

  exports.Set("Office", func);
  return exports;
}

Napi::Object Init(Napi::Env env, Napi::Object exports) {
  return Office::Init(env, exports);
}

NODE_API_MODULE(addon, Init)
