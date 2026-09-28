#!/bin/sh
# Yayın derlemesi (GitHub Pages iş akışı: .github/workflows/pages.yml): yalnızca uygulama dosyaları yayınlanır,
# README.md, CLAUDE.md ve git dosyaları değil. Uygulamanın kendisinin derleme adımı yok; bu betik yalnızca yayın içindir.
# Uygulamaya yeni bir üst düzey dosya ya da klasör eklenirse buraya da eklenmeli.
set -e
rm -rf dist
mkdir dist
cp -r index.html css js dist/
