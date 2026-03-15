cmake_minimum_required(VERSION 3.16)
project(SumColumns LANGUAGES CXX)

set(CMAKE_POSITION_INDEPENDENT_CODE ON)

set(CMAKE_AUTOMOC ON)
set(CMAKE_CXX_STANDARD 17)

# Добавляем все нужные модули Qt
find_package(Qt6 COMPONENTS Core Gui Xml REQUIRED)

# Добавь эту строку специально для поиска приватных зависимостей QXlsx
find_package(Qt6Gui REQUIRED)
find_package(Qt6GuiPrivate REQUIRED)

# 1. Указываем правильный путь к папке, где лежит CMakeLists.txt библиотеки QXlsx
add_subdirectory(3rdparty/QXlsx-1.5.0)

add_library(sum_columns SHARED
    sum_columns.cpp
    sum_columns.h
)

# 2. Пути для инклудов (теперь без лишней папки QXlsx в пути)
target_include_directories(sum_columns PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}
    ${CMAKE_CURRENT_SOURCE_DIR}/3rdparty/QXlsx-1.5.0/header
)

target_link_libraries(sum_columns PRIVATE
    Qt6::Core
    Qt6::Gui
    Qt6::Xml
    QXlsx  # Это имя цели из 3rdparty/QXlsx-1.5.0/CMakeLists.txt
)

# Настройки выходного файла
set_target_properties(sum_columns PROPERTIES
    OUTPUT_NAME "sum_columns"
    PREFIX ""
    SUFFIX ".so"
)

# Копируем манифест
add_custom_command(TARGET sum_columns POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy
            ${CMAKE_CURRENT_SOURCE_DIR}/manifest.json
            $<TARGET_FILE_DIR:sum_columns>/manifest.json
    COMMENT "Копируем manifest.json"
)
