from pathlib import Path
root=Path(__file__).resolve().parents[1]
html=(root/'dist/index.html').read_text()
html=html.replace('<link rel="stylesheet" href="style.css">','<style>'+(root/'dist/style.css').read_text()+'</style>')
html=html.replace('<script src="app.js"></script>','<script>'+(root/'dist/app.js').read_text().replace('</script','<\\/script')+'</script>')
(root/'apps-script/Index.html').write_text(html)
print('Apps-Script Index.html generated')
