[app]

title = PMonitor
package.name = pmonitor
package.domain = ua.shkarupa

source.dir = .
source.include_exts = py,png,jpg,kv,json,xlsx

version = 1.0

requirements = python3,kivy,kivymd,pandas,numpy,openpyxl

orientation = all

fullscreen = 0

icon.filename = %(source.dir)s/icon.png
presplash.filename = %(source.dir)s/presplash.png

android.api = 33
android.minapi = 24
android.archs = arm64-v8a

android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

[buildozer]

log_level = 2
warn_on_root = 1