cmake_minimum_required(VERSION 4.1)
project(Windows_Wallpaper_Changer C)

set(CMAKE_C_STANDARD 23)

add_executable(Windows_Wallpaper_Changer WIN32 main.c vdesktop.c resources.rc)

target_link_libraries(Windows_Wallpaper_Changer
        user32
        shell32
        ole32
        gdi32
        comctl32
)
