#ifndef OFFICE_H
#define OFFICE_H

#include <napi.h>
#include "LibreOfficeKit/LibreOfficeKit.h"
#include "LibreOfficeKit/LibreOfficeKitInit.h"

class Office : public Napi::ObjectWrap<Office> {
 public:
  static Napi::Object Init(Napi::Env env, Napi::Object exports);
  Office(const Napi::CallbackInfo& info);
  void Load(const Napi::CallbackInfo& info);
  void Close(const Napi::CallbackInfo& info);
  void Convert(const Napi::CallbackInfo& info);

 private:
  LibreOfficeKit* office;
  void SetOffice(LibreOfficeKit* office) { this->office = office; }
};

#endif
