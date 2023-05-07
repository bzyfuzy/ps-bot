import json
import win32com.client
ps_app = win32com.client.Dispatch("Photoshop.Application")

f = open('per_page_455.json', encoding='utf8')
data = json.load(f)


def edit_qr(doc, id, layer_group):
    qr_layer = layer_group.ArtLayers.Item("qr")
    ps_app.ActiveDocument.ActiveLayer = qr_layer
    desiredLeft = qr_layer.Bounds[0]
    desiredTop = qr_layer.Bounds[1]
    width = qr_layer.Bounds[2] - qr_layer.Bounds[0]
    height = qr_layer.Bounds[3] - qr_layer.Bounds[1]
    doc.Selection.SelectAll()
    doc.Selection.Clear()
    imageFile = fr"C:\Users\bzyfu\Documents\code\work\py\toyota\may_07_toyota\qrs\{id}.png"
    imageDoc = ps_app.Open(imageFile)
    imageDoc.ActiveLayer.Copy()
    ps_app.ActiveDocument = doc
    doc.ActiveLayer = qr_layer
    doc.Paste()
    currentWidth = qr_layer.Bounds[2] - qr_layer.Bounds[0]
    currentHeight = qr_layer.Bounds[3] - qr_layer.Bounds[1]
    scaleFactorX = width / currentWidth
    scaleFactorY = height / currentHeight
    qr_layer.Resize(scaleFactorX * 100, scaleFactorY * 100)
    currentLeft = qr_layer.Bounds[0]
    currentTop = qr_layer.Bounds[1]
    deltaLeft = desiredLeft - currentLeft
    deltaTop = desiredTop - currentTop
    qr_layer.Translate(deltaLeft, deltaTop)
    imageDoc.Close(2)
    return qr_layer


def edit_text(text, layer_group):
    text_layer = layer_group.ArtLayers.Item("text")
    textItem = text_layer.TextItem
    textItem.Contents = text


for page, page_data in enumerate(data):
    doc = ps_app.Open(
        r'C:\Users\bzyfu\Documents\code\work\py\toyota\layout.tif')
    for row, row_data in enumerate(page_data):
        for col, cell in enumerate(row_data):
            temp_qr = None
            layer_group = doc.LayerSets.Item(f"card_{row}_{col}")
            new_qr = edit_qr(doc, cell["id"], layer_group)
            if not (temp_qr is None):
                print("not none")
            edit_text(cell["name"], layer_group)
            save_path = fr'C:\Users\bzyfu\Documents\code\work\py\toyota\may_07_toyota\results\result_{page}.tif'
            tiff_options = win32com.client.Dispatch(
                "Photoshop.TiffSaveOptions")
            tiff_options.AlphaChannels = True
    doc.SaveAs(save_path, tiff_options, True)
    doc.Close(2)

f.close()
