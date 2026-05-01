/**
 * MCCM最終資料 デザイン吸い出しGAS
 *
 * 最終更新: 2026-05-01 (date '+%Y年%-m月%-d日 %H:%M' で実時刻を取得して書き換えること)
 *
 * 用途:
 *  最終資料 (株式会社MCCマネジメント御中_年次ビジネスレビュー_20260430)
 *  の P2-15 と P41-50 (計24枚) を構造化JSONで吸い出す。
 *  出力JSONを learned_layouts/p02-15_p41-50.json に保存し、
 *  デザインパターン抽出 → 再生成ヘルパー関数化に使う。
 *
 * 対象プレゼンテーション:
 *  https://docs.google.com/presentation/d/1eKkxLPQLumtMZOQqX7AwJPGWZcJ_wUXAbcfTfmVfAcQ/edit
 *
 * 実行方法:
 *  1. このファイルをそのままコピーして Apps Script エディタに貼り付け
 *  2. inspectFinalDeck() を実行
 *  3. 実行ログ (Cmd+Enter または 表示 > ログ) から最後の "===JSON_START===" 〜 "===JSON_END==="
 *     の間の文字列をすべてコピー
 *  4. learned_layouts/p02-15_p41-50.json として保存
 *
 * 注意:
 *  - 大きなプレゼンなのでログが上限を超えることがある。その場合は inspectRange() で範囲を分割実行
 *  - 画像のバイナリは取得しない (URL or contentUrl のみ)
 *  - スピーカーノートも取得対象に含める
 */

// ============================================================
// 対象プレゼンテーション
// ============================================================
var SOURCE_PRES_ID = "1eKkxLPQLumtMZOQqX7AwJPGWZcJ_wUXAbcfTfmVfAcQ";

// 対象ページ範囲 (1-indexed の人間用ページ番号)
var TARGET_RANGES = [
  { from: 2,  to: 15 },
  { from: 41, to: 50 }
];

// ============================================================
// メイン関数
// ============================================================
function inspectFinalDeck() {
  var pres = SlidesApp.openById(SOURCE_PRES_ID);
  var slides = pres.getSlides();
  var result = {
    meta: {
      presentationId: SOURCE_PRES_ID,
      presentationName: pres.getName(),
      pageWidth: pres.getPageWidth(),
      pageHeight: pres.getPageHeight(),
      totalSlides: slides.length,
      generatedAt: new Date().toISOString(),
      targetRanges: TARGET_RANGES
    },
    slides: []
  };

  // 対象ページを集める (1-indexed → 0-indexed)
  var targetIndices = [];
  TARGET_RANGES.forEach(function(range) {
    for (var i = range.from; i <= range.to; i++) {
      if (i - 1 < slides.length) targetIndices.push(i - 1);
    }
  });

  Logger.log("吸い出し対象: " + targetIndices.length + "枚 (P" +
             targetIndices.map(function(i) { return i + 1; }).join(",P") + ")");

  targetIndices.forEach(function(idx) {
    try {
      var slide = slides[idx];
      var slideData = inspectSlide(slide, idx);
      result.slides.push(slideData);
      Logger.log("✓ P" + (idx + 1) + " 完了 (要素数: " + slideData.elements.length + ")");
    } catch (e) {
      Logger.log("✗ P" + (idx + 1) + " 失敗: " + e.message);
      result.slides.push({ pageNumber: idx + 1, error: e.message });
    }
  });

  // JSON出力
  var json = JSON.stringify(result, null, 2);
  Logger.log("===JSON_START===");
  // ログは8KB制限があるので分割出力
  var chunkSize = 7000;
  for (var i = 0; i < json.length; i += chunkSize) {
    Logger.log(json.substring(i, i + chunkSize));
  }
  Logger.log("===JSON_END===");
  Logger.log("\n総文字数: " + json.length + " bytes");
  Logger.log("吸い出し完了: " + result.slides.length + "枚");

  // 大きすぎる場合は Drive に保存もする
  saveToDriveAsJson(json, "p02-15_p41-50.json");

  return result;
}

/**
 * 範囲を限定して実行する補助関数
 * 例: inspectRange(2, 15) で P2-15のみ
 */
function inspectRange(fromPage, toPage) {
  var pres = SlidesApp.openById(SOURCE_PRES_ID);
  var slides = pres.getSlides();
  var result = {
    meta: {
      presentationId: SOURCE_PRES_ID,
      presentationName: pres.getName(),
      pageWidth: pres.getPageWidth(),
      pageHeight: pres.getPageHeight(),
      generatedAt: new Date().toISOString(),
      range: { from: fromPage, to: toPage }
    },
    slides: []
  };

  for (var i = fromPage; i <= toPage; i++) {
    if (i - 1 >= slides.length) continue;
    try {
      var slide = slides[i - 1];
      var slideData = inspectSlide(slide, i - 1);
      result.slides.push(slideData);
      Logger.log("✓ P" + i + " 完了 (要素数: " + slideData.elements.length + ")");
    } catch (e) {
      Logger.log("✗ P" + i + " 失敗: " + e.message);
    }
  }

  var json = JSON.stringify(result, null, 2);
  Logger.log("===JSON_START===");
  var chunkSize = 7000;
  for (var k = 0; k < json.length; k += chunkSize) {
    Logger.log(json.substring(k, k + chunkSize));
  }
  Logger.log("===JSON_END===");

  saveToDriveAsJson(json, "p" + fromPage + "-" + toPage + "_inspect.json");
  return result;
}

// ============================================================
// スライド単位の検査
// ============================================================
function inspectSlide(slide, idx) {
  var data = {
    pageNumber: idx + 1,
    objectId: slide.getObjectId(),
    layoutName: getLayoutName(slide),
    background: getBackground(slide),
    speakerNotes: safeText(slide.getNotesPage() && slide.getNotesPage().getSpeakerNotesShape()),
    elements: []
  };

  var elements = slide.getPageElements();
  elements.forEach(function(el, elIdx) {
    try {
      data.elements.push(inspectElement(el, elIdx));
    } catch (e) {
      data.elements.push({ index: elIdx, error: e.message, type: tryGetType(el) });
    }
  });

  return data;
}

function getLayoutName(slide) {
  try {
    var layout = slide.getLayout();
    if (!layout) return null;
    return layout.getLayoutName ? layout.getLayoutName() : (layout.getObjectId ? layout.getObjectId() : null);
  } catch (e) { return null; }
}

function getBackground(slide) {
  try {
    var bg = slide.getBackground();
    var type = bg.getType();
    var info = { type: type ? type.toString() : null };
    if (type === SlidesApp.PageBackgroundType.SOLID) {
      var fill = bg.getSolidFill();
      if (fill) {
        info.color = colorToHex(fill.getColor());
        info.alpha = fill.getAlpha();
      }
    }
    return info;
  } catch (e) { return { error: e.message }; }
}

// ============================================================
// 要素単位の検査
// ============================================================
function inspectElement(el, idx) {
  var type = el.getPageElementType();
  var data = {
    index: idx,
    objectId: el.getObjectId(),
    type: type.toString(),
    geometry: {
      x: round(el.getLeft()),
      y: round(el.getTop()),
      w: round(el.getWidth()),
      h: round(el.getHeight()),
      rotation: round(el.getRotation())
    },
    title: el.getTitle ? safeCall(function() { return el.getTitle(); }) : null,
    description: el.getDescription ? safeCall(function() { return el.getDescription(); }) : null
  };

  switch (type) {
    case SlidesApp.PageElementType.SHAPE:
      data.shape = inspectShape(el.asShape());
      break;
    case SlidesApp.PageElementType.IMAGE:
      data.image = inspectImage(el.asImage());
      break;
    case SlidesApp.PageElementType.LINE:
      data.line = inspectLine(el.asLine());
      break;
    case SlidesApp.PageElementType.TABLE:
      data.table = inspectTable(el.asTable());
      break;
    case SlidesApp.PageElementType.GROUP:
      data.group = inspectGroup(el.asGroup());
      break;
    case SlidesApp.PageElementType.WORD_ART:
      data.wordArt = { text: safeCall(function() { return el.asWordArt().getRenderedText(); }) };
      break;
    case SlidesApp.PageElementType.SHEETS_CHART:
      data.sheetsChart = { chartId: safeCall(function() { return el.asSheetsChart().getChartId(); }) };
      break;
    case SlidesApp.PageElementType.VIDEO:
      data.video = { url: safeCall(function() { return el.asVideo().getUrl(); }) };
      break;
    case SlidesApp.PageElementType.SPEAKER_SPOTLIGHT:
      data.speakerSpotlight = {};
      break;
    default:
      // unknown
      break;
  }

  return data;
}

// ============================================================
// SHAPE の検査
// ============================================================
function inspectShape(shape) {
  var info = {
    shapeType: shape.getShapeType() ? shape.getShapeType().toString() : null,
    placeholderType: null,
    placeholderIndex: null,
    contentAlignment: safeCall(function() {
      var ca = shape.getContentAlignment();
      return ca ? ca.toString() : null;
    }),
    autofit: safeCall(function() {
      var af = shape.getAutofit && shape.getAutofit();
      if (!af) return null;
      return {
        autofitType: af.getAutofitType ? af.getAutofitType().toString() : null,
        fontScale: af.getFontScale ? af.getFontScale() : null,
        lineSpacingReduction: af.getLineSpacingReduction ? af.getLineSpacingReduction() : null
      };
    }),
    fill: getShapeFill(shape),
    border: getShapeBorder(shape),
    text: inspectTextRange(shape.getText())
  };

  // プレースホルダー情報
  try {
    var pt = shape.getPlaceholderType();
    if (pt) info.placeholderType = pt.toString();
    info.placeholderIndex = shape.getPlaceholderIndex();
  } catch (e) {}

  return info;
}

function getShapeFill(shape) {
  try {
    var fill = shape.getFill();
    if (!fill) return null;
    var type = fill.getType();
    var info = { type: type ? type.toString() : null };
    if (type === SlidesApp.FillType.SOLID) {
      var sf = fill.getSolidFill();
      if (sf) {
        info.color = colorToHex(sf.getColor());
        info.alpha = sf.getAlpha();
      }
    }
    return info;
  } catch (e) { return { error: e.message }; }
}

function getShapeBorder(shape) {
  try {
    var border = shape.getBorder();
    if (!border) return null;
    var info = {
      isVisible: border.isVisible ? border.isVisible() : null,
      weight: border.getWeight ? border.getWeight() : null,
      dashStyle: border.getDashStyle ? (border.getDashStyle() ? border.getDashStyle().toString() : null) : null
    };
    var lineFill = border.getLineFill && border.getLineFill();
    if (lineFill) {
      var ft = lineFill.getFillType ? lineFill.getFillType() : null;
      info.lineFillType = ft ? ft.toString() : null;
      if (ft === SlidesApp.LineFillType.SOLID) {
        var sf = lineFill.getSolidFill();
        if (sf) info.color = colorToHex(sf.getColor());
      }
    }
    return info;
  } catch (e) { return { error: e.message }; }
}

// ============================================================
// IMAGE / LINE / TABLE / GROUP の検査
// ============================================================
function inspectImage(img) {
  return {
    contentUrl: safeCall(function() { return img.getContentUrl(); }),
    sourceUrl: safeCall(function() { return img.getSourceUrl(); }),
    altText: safeCall(function() { return img.getDescription(); }),
    border: getShapeBorder(img),
    brightness: safeCall(function() { return img.getBrightness && img.getBrightness(); }),
    contrast: safeCall(function() { return img.getContrast && img.getContrast(); }),
    transparency: safeCall(function() { return img.getTransparency && img.getTransparency(); })
  };
}

function inspectLine(line) {
  return {
    lineType: safeCall(function() { return line.getLineType() ? line.getLineType().toString() : null; }),
    lineCategory: safeCall(function() { return line.getLineCategory && line.getLineCategory() ? line.getLineCategory().toString() : null; }),
    weight: safeCall(function() { return line.getWeight(); }),
    dashStyle: safeCall(function() { return line.getDashStyle() ? line.getDashStyle().toString() : null; }),
    startArrow: safeCall(function() { return line.getStartArrow() ? line.getStartArrow().toString() : null; }),
    endArrow: safeCall(function() { return line.getEndArrow() ? line.getEndArrow().toString() : null; }),
    color: safeCall(function() {
      var lf = line.getLineFill();
      if (!lf) return null;
      if (lf.getFillType() === SlidesApp.LineFillType.SOLID) {
        return colorToHex(lf.getSolidFill().getColor());
      }
      return null;
    })
  };
}

function inspectTable(table) {
  var rows = table.getNumRows();
  var cols = table.getNumColumns();
  var data = {
    numRows: rows,
    numColumns: cols,
    rowHeights: [],
    columnWidths: [],
    cells: []
  };
  for (var r = 0; r < rows; r++) {
    data.rowHeights.push(safeCall(function() { return table.getRow(r).getMinimumHeight(); }));
  }
  for (var c = 0; c < cols; c++) {
    data.columnWidths.push(safeCall(function() { return table.getColumn(c).getWidth(); }));
  }
  for (var r2 = 0; r2 < rows; r2++) {
    for (var c2 = 0; c2 < cols; c2++) {
      try {
        var cell = table.getCell(r2, c2);
        data.cells.push({
          row: r2,
          col: c2,
          rowSpan: cell.getRowSpan ? cell.getRowSpan() : 1,
          colSpan: cell.getColumnSpan ? cell.getColumnSpan() : 1,
          fill: getShapeFill(cell),
          contentAlignment: safeCall(function() { return cell.getContentAlignment() ? cell.getContentAlignment().toString() : null; }),
          text: inspectTextRange(cell.getText())
        });
      } catch (e) {
        data.cells.push({ row: r2, col: c2, error: e.message });
      }
    }
  }
  return data;
}

function inspectGroup(group) {
  var children = group.getChildren();
  var info = { childCount: children.length, children: [] };
  children.forEach(function(child, i) {
    try {
      info.children.push(inspectElement(child, i));
    } catch (e) {
      info.children.push({ index: i, error: e.message });
    }
  });
  return info;
}

// ============================================================
// TEXT の検査 (段落 + ラン)
// ============================================================
function inspectTextRange(textRange) {
  if (!textRange) return null;
  var text = "";
  try { text = textRange.asString(); } catch (e) {}
  var info = {
    text: text,
    paragraphs: [],
    runs: []
  };

  // 段落単位
  try {
    var paragraphs = textRange.getParagraphs();
    paragraphs.forEach(function(p, i) {
      try {
        var pStyle = p.getRange().getParagraphStyle();
        info.paragraphs.push({
          index: i,
          text: p.getRange().asString(),
          alignment: safeCall(function() { return pStyle.getParagraphAlignment() ? pStyle.getParagraphAlignment().toString() : null; }),
          indentStart: safeCall(function() { return pStyle.getIndentStart && pStyle.getIndentStart(); }),
          indentEnd: safeCall(function() { return pStyle.getIndentEnd && pStyle.getIndentEnd(); }),
          indentFirstLine: safeCall(function() { return pStyle.getIndentFirstLine && pStyle.getIndentFirstLine(); }),
          lineSpacing: safeCall(function() { return pStyle.getLineSpacing && pStyle.getLineSpacing(); }),
          spaceAbove: safeCall(function() { return pStyle.getSpaceAbove && pStyle.getSpaceAbove(); }),
          spaceBelow: safeCall(function() { return pStyle.getSpaceBelow && pStyle.getSpaceBelow(); }),
          spacingMode: safeCall(function() { return pStyle.getSpacingMode && pStyle.getSpacingMode() ? pStyle.getSpacingMode().toString() : null; }),
          direction: safeCall(function() { return pStyle.getTextDirection && pStyle.getTextDirection() ? pStyle.getTextDirection().toString() : null; })
        });
      } catch (e) {
        info.paragraphs.push({ index: i, error: e.message });
      }
    });
  } catch (e) {}

  // ラン単位 (フォント・色などの実値)
  try {
    var runs = textRange.getRuns();
    runs.forEach(function(run, i) {
      try {
        var s = run.getTextStyle();
        info.runs.push({
          index: i,
          text: run.asString(),
          fontFamily: safeCall(function() { return s.getFontFamily(); }),
          fontSize: safeCall(function() { return s.getFontSize(); }),
          bold: safeCall(function() { return s.isBold(); }),
          italic: safeCall(function() { return s.isItalic(); }),
          underline: safeCall(function() { return s.isUnderline(); }),
          strikethrough: safeCall(function() { return s.isStrikethrough(); }),
          smallCaps: safeCall(function() { return s.isSmallCaps && s.isSmallCaps(); }),
          baselineOffset: safeCall(function() { return s.getBaselineOffset() ? s.getBaselineOffset().toString() : null; }),
          foregroundColor: safeCall(function() {
            var fg = s.getForegroundColor();
            return fg ? colorToHex(fg) : null;
          }),
          backgroundColor: safeCall(function() {
            var bg = s.getBackgroundColor();
            try {
              var c = bg.asRgbColor();
              return c ? "#" + pad(c.asHexString().substring(1)) : null;
            } catch (e) { return null; }
          }),
          link: safeCall(function() {
            var link = s.getLink && s.getLink();
            if (!link) return null;
            return {
              linkType: link.getLinkType ? link.getLinkType().toString() : null,
              url: link.getUrl ? link.getUrl() : null
            };
          })
        });
      } catch (e) {
        info.runs.push({ index: i, error: e.message });
      }
    });
  } catch (e) {}

  // リスト書式
  try {
    var listStyles = [];
    textRange.getParagraphs().forEach(function(p, i) {
      try {
        var ls = p.getRange().getListStyle();
        if (ls && ls.isInList && ls.isInList()) {
          listStyles.push({
            paragraphIndex: i,
            glyph: safeCall(function() { return ls.getGlyph(); }),
            nestingLevel: safeCall(function() { return ls.getNestingLevel(); })
          });
        }
      } catch (e) {}
    });
    if (listStyles.length > 0) info.listStyles = listStyles;
  } catch (e) {}

  return info;
}

// ============================================================
// ユーティリティ
// ============================================================
function colorToHex(color) {
  if (!color) return null;
  try {
    var ct = color.getColorType();
    if (ct === SlidesApp.ColorType.RGB) {
      return color.asRgbColor().asHexString();
    } else if (ct === SlidesApp.ColorType.THEME) {
      return "THEME:" + color.asThemeColor().getThemeColorType().toString();
    }
    return ct ? ct.toString() : null;
  } catch (e) {
    return null;
  }
}

function safeCall(fn) {
  try { return fn(); } catch (e) { return null; }
}

function safeText(shape) {
  try { return shape && shape.getText() ? shape.getText().asString() : null; } catch (e) { return null; }
}

function tryGetType(el) {
  try { return el.getPageElementType().toString(); } catch (e) { return "UNKNOWN"; }
}

function round(v) {
  if (v === null || v === undefined || isNaN(v)) return v;
  return Math.round(v * 100) / 100;
}

function pad(s) {
  while (s.length < 6) s = "0" + s;
  return s;
}

// ============================================================
// Drive にJSONとして保存 (ログ上限対策)
// ============================================================
function saveToDriveAsJson(json, filename) {
  try {
    var folder;
    var folderName = "MCCM_inspect_output";
    var folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }
    var file = folder.createFile(filename, json, MimeType.PLAIN_TEXT);
    Logger.log("\n✓ Driveに保存しました: " + folder.getName() + "/" + filename);
    Logger.log("  ファイルID: " + file.getId());
    Logger.log("  URL: " + file.getUrl());
  } catch (e) {
    Logger.log("\n✗ Drive保存失敗: " + e.message);
  }
}
