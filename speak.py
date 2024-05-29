import sys
import win32com.client
import ctypes

INFINITE = ctypes.c_ulong(-1)


def get_voices(sapi):
    voices = []
    for voice in sapi.GetVoices():
        name = voice.GetAttribute("Name")
        voices.append(name)

    return voices


def find_voice(sapi, voice_name):
    for voice in sapi.GetVoices():
        name = voice.GetAttribute("Name")
        # print(f"Name={name}, Id={voice.Id}")
        if name == voice_name:
            return voice

    return None


def main():
    sapi = win32com.client.Dispatch("SAPI.SpVoice")

    if len(sys.argv) == 1:
        print(f"Usage: {sys.argv[0]} voice")
        print("\nVoices:")
        for name in get_voices(sapi):
            print(f"    {name}")
        return

    voice_name = sys.argv[1]

    voice = find_voice(sapi, voice_name)
    if not voice:
        print("voice `{voice_name}` not found")
        return

    sapi.Voice = voice

    sapi.Speak("Open the pod-bay doors, Hal.", 1)
    # print("Waiting until done")
    # sapi.WaitUntilDone(-1)

    sapi.Speak("Open the pod-bay doors, Hal.", 1)
    # print("Waiting until done")
    # sapi.WaitUntilDone(-1)

    print("Success!")


main()
