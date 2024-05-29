"""sapi_tts_engine.py: Proof-of-concept implementation of SAPI TTS Engine"""

__author__     = "Bostjan Vesnicer"
__copyright__  = "Copyright (c) 2022, Bostjan Vesnicer"
__license__    = "SPDX-License-Identifier: BSD-2-Clause"
__version__    = "0.9.0"
__maintainer__ = "Bostjan Vesnicer"
__email__      = "bostjan.vesnicer@gmail.com"
__status__     = "Prototype"

import comtypes
import winreg
import ctypes
import comtypes.server.localserver
from comtypes.hresult import S_OK

# from comtypes.client import GetModule
# GetModule("pysapi.tlb")
from comtypes.gen.PySAPILib import (
    TTSEngine, WAVEFORMATEX, SPVTEXTFRAG, SPEVENT
)

# helper functions

def make_wave_format():
    wave_format = WAVEFORMATEX()
    wave_format.wFormatTag = WAVE_FORMAT_PCM
    wave_format.nChannels = 1
    wave_format.nSamplesPerSec = 16000
    wave_format.nAvgBytesPerSec = 32000
    wave_format.nBlockAlign = 2
    wave_format.wBitsPerSample = 16
    wave_format.cbSize = 0
    return wave_format


def print_wave_format(wave_format):
    print(f"wFormatTag={wave_format.wFormatTag}"
          f", nSamplesPerSec={wave_format.nSamplesPerSec}"
          f", nChannels={wave_format.nChannels}"
          f", nAvgBytesPerSec={wave_format.nAvgBytesPerSec}"
          f", nBlockAlign={wave_format.nBlockAlign}"
          f", wBitsPerSample={wave_format.wBitsPerSample}"
          f", cbSize={wave_format.cbSize}")


def load_wav(filename):
    with open(filename, "rb") as f:
        data = f.read()
    # skip header
    data = data[44:]
    return ctypes.create_string_buffer(data, len(data))


# constants
SPDFID_WaveFormatEx = comtypes.GUID("{C31ADBAE-527F-4ff5-A230-F62BB61FF70C}")
WAVE_FORMAT_PCM = 1
SPET_LPARAM_IS_UNDEFINED = 0
SPEI_SENTENCE_BOUNDARY = 7


class TTSEngineImpl(TTSEngine):
    # _com_interfaces_ = [ISpTTSEngine, ISpObjectWithToken]
    _reg_threading_ = "Both"
    _reg_progid_ = "PySAPILib.TTSEngine.1"
    _reg_novers_progid_ = "PySAPILib.TTSEngine"
    _reg_desc_ = "Python COM Text-To-Speech Engine"
    _reg_clsctx_ = comtypes.CLSCTX_INPROC_SERVER | comtypes.CLSCTX_LOCAL_SERVER
    _regcls_ = comtypes.server.localserver.REGCLS_MULTIPLEUSE
    _reg_clsid_ = "{4DFFD59B-4DF3-4366-B053-DFF9BE002EFB}"

    def __init__(self):
        print(f"Initializing {__class__._reg_desc_}\n")
        self.data = load_wav("hal.wav")
        ctypes.windll.ole32.CoTaskMemAlloc.argtypes = [ctypes.c_size_t]
        ctypes.windll.ole32.CoTaskMemAlloc.restype = ctypes.c_void_p

    def ISpTTSEngine_Speak(self, this, speak_flags, format_id, wave_format, text_frag_list, output_site):
        """ HRESULT Speak
                DWORD                 dwSpeakFlags,
                REFGUID               rguidFormatId,
                const WaveFormatEx   *pWaveFormatEx,
                const SPVTEXTFRAG    *pTextFragList,
                ISpTTSEngineSite     *pOutputSite
        """

        print(f"\nSpeak")

        print(f"speak_flags={speak_flags}")

        text_frag = text_frag_list
        text_len = 0

        while text_frag:
            text_frag = text_frag[0]
            print(f"pTextStart={text_frag.pTextStart}"
                  f", ulTextLen={text_frag.ulTextLen}"
                  f", ulTextSrcOffset={text_frag.ulTextSrcOffset}")
            text_len += text_frag.ulTextLen
            text_frag = text_frag.pNext

        # FIXME: handle actions
        actions = output_site.GetActions()
        print(f"actions={actions}")

        # FIXME: properly handle events
        event = SPEVENT()
        event.eEventId = SPEI_SENTENCE_BOUNDARY
        event.elParamType = SPET_LPARAM_IS_UNDEFINED
        event.ulStreamNum = 0
        event.ullAudioStreamOffset = 0
        event.lParam = 0
        event.wParam = text_len

        result = output_site.AddEvents(ctypes.pointer(event), 1)
        if result != S_OK:
            raise RuntimeError("AddEvents failed")

        # write speech data
        num_bytes = len(self.data)
        bytes_written = output_site.Write(self.data, num_bytes)
        if bytes_written != num_bytes:
            print("bytes_written != num_bytes")

        print("Speak END")

        return S_OK

    def ISpTTSEngine_GetOutputFormat(self, this, target_fmt_id, target_wave_format, output_format_id, output_wave_format):
        """ HRESULT GetOutputFormat
                const GUID          *pTargetFmtId
                const WAVEFORMATEX  *pTargetWaveFormatEx
                GUID                *pOutputFormatId
                WAVEFORMATEX       **ppCoMemOutputWaveFormatEx
        """

        print(f"\nGetOutputFormat")

        if not target_fmt_id:
            output_format_id[0] = SPDFID_WaveFormatEx
            buffer = ctypes.windll.ole32.CoTaskMemAlloc(ctypes.sizeof(WAVEFORMATEX))
            ctypes.memmove(buffer, ctypes.pointer(make_wave_format()), ctypes.sizeof(WAVEFORMATEX))
            wave_format=ctypes.cast(buffer, ctypes.POINTER(WAVEFORMATEX))
            output_wave_format[0] = wave_format
        else:
            print(f"target_fmt_id={target_fmt_id[0]}")
            print_wave_format(target_wave_format[0])

            output_format_id[0]=SPDFID_WaveFormatEx
            buffer = ctypes.windll.ole32.CoTaskMemAlloc(ctypes.sizeof(WAVEFORMATEX))
            ctypes.memmove(buffer, ctypes.pointer(make_wave_format()), ctypes.sizeof(WAVEFORMATEX))
            wave_format = ctypes.cast(buffer, ctypes.POINTER(WAVEFORMATEX))
            output_wave_format[0] = wave_format

        print("GetOutputFormat END")

        return S_OK

    def ISpObjectWithToken_SetObjectToken(self, this, object_token):
        """
        HRESULT SetObjectToken
            ISpObjectToken  *pToken
        """

        print(f"\nSetObjectToken")
        self.object_token = object_token

        print("SetObjectToken END")

        return S_OK

    def ISpObjectWithToken_GetObjectToken(self, this, object_token):
        """
        HRESULT GetObjectToken
            ISpObjectToken   **ppToken
        """

        print(f"\nGetObjectToken")
        object_token[0] = self.object_token

        print(f"GetObjectToken END")

        return S_OK


def register_voice(clsid, id, name, language, gender, vendor):
    # Paths for 64-bit and 32-bit applications
    paths = [
        f"SOFTWARE\\Microsoft\\Speech\\Voices\\Tokens\\{id}",
        f"SOFTWARE\\WOW6432Node\\Microsoft\\Speech\\Voices\\Tokens\\{id}"
    ]

    for path in paths:
        # Create the registry key for the voice
        key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, path)

        # Set the attributes of the voice
        winreg.SetValueEx(key, None, 0, winreg.REG_SZ, name)
        winreg.SetValueEx(key, "409", 0, winreg.REG_SZ, name)
        winreg.SetValueEx(key, "CLSID", 0, winreg.REG_SZ, clsid)

        # Create the registry key for the attributes
        attr_key = winreg.CreateKey(key, "Attributes")
        winreg.SetValueEx(attr_key, "Language", 0, winreg.REG_SZ, language)
        winreg.SetValueEx(attr_key, "Gender", 0, winreg.REG_SZ, gender)
        winreg.SetValueEx(attr_key, "Vendor", 0, winreg.REG_SZ, vendor)
        winreg.SetValueEx(attr_key, "Name", 0, winreg.REG_SZ, name)

        # Close the registry keys
        winreg.CloseKey(attr_key)
        winreg.CloseKey(key)


if __name__ == "__main__":
    from comtypes.server.register import UseCommandLine
    UseCommandLine(TTSEngineImpl)

    register_voice(TTSEngineImpl._reg_clsid_, "PySAPIPoCVoice",
                   "Python SAPI PoC Voice", "409", "male", "Python")
