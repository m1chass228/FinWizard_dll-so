cmake_minimum_required(VERSION 3.16)
project(Calc_salary_birulevo LANGUAGES CXX)

set(CMAKE_POSITION_INDEPENDENT_CODE ON)

set(CMAKE_AUTOMOC ON)
set(CMAKE_CXX_STANDARD 17)

# Добавляем все нужные модули Qt
find_package(Qt6 COMPONENTS Core Gui Xml REQUIRED)
find_package(Qt6GuiPrivate REQUIRED)

# --- ВАЖНО: Впаиваем QXlsx прямо внутрь плагина ---
# Отключаем сборку QXlsx как отдельной DLL
set(BUILD_SHARED_LIBS OFF CACHE BOOL "" FORCE)
add_subdirectory(3rdparty/QXlsx-1.5.0)
# Возвращаем настройку обратно для нашего плагина
set(BUILD_SHARED_LIBS ON CACHE BOOL "" FORCE)
# --------------------------------------------------

add_library(calc_salary SHARED
    calc_salary_birulevo.cpp
    calc_salary_birulevo.h
)

target_include_directories(calc_salary PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}
    ${CMAKE_CURRENT_SOURCE_DIR}/3rdparty/QXlsx-1.5.0/header
)

target_link_libraries(calc_salary PRIVATE
    Qt6::Core
    Qt6::Gui
    Qt6::Xml
    QXlsx
)

# Настройки выходного файла (убираем lib и .so, чтобы было кроссплатформенно)
set_target_properties(calc_salary PROPERTIES
    OUTPUT_NAME "calc_salary"
    PREFIX ""
)

# Копируем манифест
add_custom_command(TARGET calc_salary POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy
            ${CMAKE_CURRENT_SOURCE_DIR}/manifest.json
            $<TARGET_FILE_DIR:calc_salary>/manifest.json
    COMMENT "Копируем manifest.json"
)
