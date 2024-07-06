import time
import psutil
import win32gui
import win32process
from ctypes import cast, POINTER
from comtypes import CoInitialize, CoUninitialize, CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


def get_spotify_processes():
    proceses = []

    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == 'Spotify.exe':
            proceses.append(proc)
    return proceses


def get_hwnds_for_pid(pid):
    # Get the window hwnd associated to the Spotify process with pid
    def callback(hwnd, hwnds):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            if found_pid == pid:
                hwnds.append(hwnd)
        return True

    hwnds = []
    win32gui.EnumWindows(callback, hwnds)
    return hwnds # Only the first item (must have only one)


def get_window_title(processes):
    for process in processes: # Find what proccess has a window with title and return then 
        try:
            hwnds = get_hwnds_for_pid(process.pid)
            if hwnds:
                hwnd = hwnds[0]
                title = win32gui.GetWindowText(hwnd)
                return title
        except Exception as e:
            print(f'Error al obtener la ventana: {e}')
            return None


def is_ad_playing():
    processes = get_spotify_processes()
    
    if processes:
        title = get_window_title(processes)
        if title:
            whitelist = ['Spotify Free']
            blacklist = ['Advertisement', 'Spotify']
            # Check the title, when ads is playing, the window title is 'Advertisement' or 'Spotify'
            # -> Alternative | When playing a song, the title is like: 'author - song' so we can filter with the '-' and knowing that when music stopped the title is 'Spotify Free'
            if (title in blacklist) or ('-' not in title and title not in whitelist):
                return True, title # Ad is playing
        
        return False, title
                  

def set_volume(volume_level):
    CoInitialize()
    try: 
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevelScalar(volume_level, None)
    finally:
        CoUninitialize()

def set_default_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    
    return volume.GetMasterVolumeLevelScalar()
    

def main():
    default_volume = 0.5    # Default volume
    
    ad_volume = 0.04        # Ads volume
    check_interval = 3

    last_status = None
    last_title = None

    try:
        while True:
            try:
                # Get the current system volume and set a new default volume
                new_volume = set_default_volume()
                if new_volume != default_volume and last_status is not True:
                    default_volume = new_volume
                    print(f'Nuevo volumen establecido al {int(default_volume*100)} %')
            except:
                print('Error al comprobar el volumen.')
                
            ad_playing, title = is_ad_playing()
            
            if ad_playing:
                if last_status is not True:
                    set_volume(ad_volume)
                    last_status = True
                if title != last_title:
                    print(f'Está sonando un anuncio. Bajando el volumen al {int(ad_volume*100)} %')   
            else:
                if last_status is not False:  
                    set_volume(default_volume)
                    print(f'Volumen establecido al {int(default_volume*100)} %')
                    last_status = False
                if title != last_title:
                    if title == 'Spotify Free':
                        print('Se ha parado la reproducción.')
                    else:
                        print(f'Está sonando: {title}')
                    
            
            last_title = title
            
            time.sleep(check_interval)
    except KeyboardInterrupt:
        print("Script terminado por el usuario.")
    except Exception as e:
        print(f"Ocurrió un error: {e}")


if __name__ == "__main__":
    main()
