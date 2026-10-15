//! Parts the model does not interpret survive a round trip: they are written back under their
//! names, with their content types and the relationships that pointed at them.

use std::collections::BTreeSet;
use std::io::{Cursor, Read, Write};

use deckcraft_model::{KeptTarget, ShapeKind, defaults};

const RELS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const RT: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn png() -> Vec<u8> {
    let img = image::RgbImage::from_pixel(4, 4, image::Rgb([10, 200, 30]));
    let mut out = Cursor::new(vec![]);
    image::DynamicImage::ImageRgb8(img).write_to(&mut out, image::ImageFormat::Png).expect("png");
    out.into_inner()
}

/// The names of a package's files.
fn names(bytes: &[u8]) -> BTreeSet<String> {
    let mut z = zip::ZipArchive::new(Cursor::new(bytes)).expect("zip");
    (0..z.len()).map(|i| z.by_index(i).expect("entry").name().to_string()).filter(|n| !n.ends_with('/')).collect()
}

fn part(bytes: &[u8], name: &str) -> String {
    let mut z = zip::ZipArchive::new(Cursor::new(bytes)).expect("zip");
    let mut f = z.by_name(name).unwrap_or_else(|e| panic!("{name}: {e}"));
    let mut s = String::new();
    f.read_to_string(&mut s).expect("utf8");
    s
}

/// A fresh deck's package with extra parts, and edits to its XML parts.
fn with_parts(parts: &[(&str, Vec<u8>)], edit: impl Fn(&str, String) -> String) -> Vec<u8> {
    let base = deckcraft_pptx::export(&defaults::new_presentation(None)).expect("export");
    let mut z = zip::ZipArchive::new(Cursor::new(base)).expect("zip");
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let o = zip::write::SimpleFileOptions::default();
    for i in 0..z.len() {
        let mut f = z.by_index(i).expect("entry");
        let name = f.name().to_string();
        let mut d = vec![];
        f.read_to_end(&mut d).expect("read");
        if name.ends_with(".xml") || name.ends_with(".rels") {
            d = edit(&name, String::from_utf8(d).expect("utf8")).into_bytes();
        }
        out.start_file(name, o).expect("start");
        out.write_all(&d).expect("write");
    }
    for (name, data) in parts {
        out.start_file(*name, o).expect("start");
        out.write_all(data).expect("write");
    }
    out.finish().expect("finish").into_inner()
}

const CHART: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="10"/><c:axId val="20"/></c:barChart><c:catAx><c:axId val="10"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:crossAx val="10"/></c:valAx></c:plotArea><c:legend><c:legendPos val="b"/></c:legend></c:chart><c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>"#;

/// A deck with everything the model skips: a thumbnail, printer settings, a chart with its
/// embedded workbook, an OLE object with a preview picture, and an orphan diagram data part.
fn deck() -> Vec<u8> {
    let frames = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="900" name="Chart 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="901" name="Object 2"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="914400"/><a:ext cx="1828800" cy="914400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj name="Sheet" r:id="rIdOle" imgW="100" imgH="100" progId="Excel.Sheet.12"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdPic"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>"#;
    let chart_rels = format!(
        r#"<Relationships xmlns="{RELS_NS}"><Relationship Id="rId1" Type="{RT}/package" Target="../embeddings/Microsoft_Excel_Sheet1.xlsx"/></Relationships>"#
    );
    with_parts(
        &[
            ("docProps/thumbnail.jpeg", vec![0xFF, 0xD8, 0xFF, 0xE0, 1, 2, 3]),
            ("ppt/printerSettings/printerSettings1.bin", vec![7; 32]),
            ("ppt/charts/chart1.xml", CHART.as_bytes().to_vec()),
            ("ppt/charts/_rels/chart1.xml.rels", chart_rels.into_bytes()),
            ("ppt/embeddings/Microsoft_Excel_Sheet1.xlsx", b"PK\x03\x04not-really-a-workbook".to_vec()),
            ("ppt/embeddings/oleObject1.bin", vec![0xD0, 0xCF, 0x11, 0xE0, 9, 9]),
            ("ppt/media/preview.png", png()),
            ("ppt/diagrams/data1.xml", br#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>"#.to_vec()),
            // A handout master with its own theme: the export regenerates theme parts, so the
            // master's theme relationship cannot be restored and the master is left out whole.
            (
                "ppt/handoutMasters/handoutMaster1.xml",
                br#"<p:handoutMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree/></p:cSld></p:handoutMaster>"#.to_vec(),
            ),
            (
                "ppt/handoutMasters/_rels/handoutMaster1.xml.rels",
                format!(r#"<Relationships xmlns="{RELS_NS}"><Relationship Id="rId1" Type="{RT}/theme" Target="../theme/theme2.xml"/></Relationships>"#).into_bytes(),
            ),
            ("ppt/theme/theme2.xml", br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Handout"/>"#.to_vec()),
        ],
        |name, s| {
            match name {
            "[Content_Types].xml" => s.replacen(
                "</Types>",
                r#"<Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/><Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/><Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.presentationml.printerSettings"/><Override PartName="/ppt/embeddings/oleObject1.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/><Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/ppt/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/></Types>"#,
                1,
            ),
            "_rels/.rels" => s.replacen(
                "</Relationships>",
                r#"<Relationship Id="rIdThumb" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/></Relationships>"#,
                1,
            ),
            "ppt/_rels/presentation.xml.rels" => s.replacen(
                "</Relationships>",
                &format!(
                    r#"<Relationship Id="rIdPrn" Type="{RT}/printerSettings" Target="printerSettings/printerSettings1.bin"/><Relationship Id="rIdHo" Type="{RT}/handoutMaster" Target="handoutMasters/handoutMaster1.xml"/></Relationships>"#
                ),
                1,
            ),
            "ppt/slides/slide1.xml" => s.replacen("</p:spTree>", &format!("{frames}</p:spTree>"), 1),
            "ppt/slides/_rels/slide1.xml.rels" => s.replacen(
                "</Relationships>",
                &format!(
                    r#"<Relationship Id="rIdChart" Type="{RT}/chart" Target="../charts/chart1.xml"/><Relationship Id="rIdOle" Type="{RT}/oleObject" Target="../embeddings/oleObject1.bin"/><Relationship Id="rIdPic" Type="{RT}/image" Target="../media/preview.png"/></Relationships>"#
                ),
                1,
            ),
            _ => s,
        }
        },
    )
}

/// A part's content type by `[Content_Types].xml`: its override, else its extension's default.
fn content_type(ct: &str, part_name: &str) -> String {
    if let Some(i) = ct.find(&format!(r#"PartName="{part_name}""#)) {
        return attr(&ct[i..], "PartName", "ContentType").to_string();
    }
    let ext = part_name.rsplit('.').next().expect("extension");
    let i = ct.find(&format!(r#"Extension="{ext}""#)).unwrap_or_else(|| panic!("no content type for {part_name}: {ct}"));
    attr(&ct[i..], "Extension", "ContentType").to_string()
}

fn pres_rels_has_handout(out: &[u8]) -> bool {
    part(out, "ppt/_rels/presentation.xml.rels").contains("handoutMaster")
}

fn rel_target(rels: &str, id: &str) -> String {
    let i = rels.find(&format!(r#"Id="{id}""#)).unwrap_or_else(|| panic!("no relationship {id} in {rels}"));
    let rest = &rels[i..];
    let t = rest.find("Target=\"").map(|j| &rest[j + 8..]).expect("target");
    t[..t.find('"').expect("quote")].to_string()
}

fn attr<'a>(xml: &'a str, after: &str, name: &str) -> &'a str {
    let i = xml.find(after).unwrap_or_else(|| panic!("no {after}"));
    let rest = &xml[i..];
    let j = rest.find(&format!("{name}=\"")).expect("attr");
    let v = &rest[j + name.len() + 2..];
    &v[..v.find('"').expect("quote")]
}

#[test]
fn unmodelled_parts_come_back_with_their_relationships() {
    let src = deck();
    let p = deckcraft_pptx::import(&src).expect("import");
    let kept: BTreeSet<&str> = p.kept_parts.iter().map(|k| k.name.as_str()).collect();
    assert_eq!(
        kept,
        [
            "docProps/thumbnail.jpeg",
            "ppt/diagrams/data1.xml",
            "ppt/embeddings/Microsoft_Excel_Sheet1.xlsx",
            "ppt/embeddings/oleObject1.bin",
            "ppt/printerSettings/printerSettings1.bin"
        ]
        .into_iter()
        .collect(),
        "everything the model did not consume, orphans included; the chart and the preview picture are modelled, the handout master cannot be written whole"
    );
    assert_eq!(
        p.kept_parts.iter().find(|k| k.name.ends_with("oleObject1.bin")).map(|k| k.content_type.as_str()),
        Some("application/vnd.openxmlformats-officedocument.oleObject")
    );
    assert_eq!(p.kept_package_rels.len(), 1, "the thumbnail; core and app properties are regenerated");
    assert_eq!(p.kept_rels.len(), 1, "printer settings; masters, slides, theme… are regenerated: {:?}", p.kept_rels);
    assert!(p.slides[0].kept_rels.iter().all(|r| matches!(r.target, KeptTarget::Part { .. })), "{:?}", p.slides[0].kept_rels);
    let slide = &p.slides[0];
    let ole = slide.shapes.iter().find(|s| s.name == "Object 2").expect("the OLE frame");
    let ShapeKind::Opaque { rels, preview, .. } = &ole.kind else { panic!("{:?}", ole.kind) };
    assert!(preview.is_some(), "the preview picture is a media item");
    assert!(rels.iter().any(|r| r.id == "rIdOle" && matches!(&r.target, KeptTarget::Part { name } if name == "ppt/embeddings/oleObject1.bin")));
    assert!(rels.iter().any(|r| r.id == "rIdPic" && matches!(r.target, KeptTarget::Media { .. })), "the picture is not stored twice");
    let chart = slide.shapes.iter().find_map(|s| match &s.kind {
        ShapeKind::Chart(c) => Some(c),
        _ => None,
    });
    assert_eq!(chart.map(|c| c.kept_rels.len()), Some(1), "the chart keeps its workbook relationship");

    let out = deckcraft_pptx::export(&p).expect("export");
    let (a, b) = (names(&src), names(&out));
    // Media items are the model's own and get the export's names; the handout master and its
    // theme are left out (see `deck`); everything else keeps its name.
    let lost: Vec<_> =
        a.difference(&b).filter(|n| !n.starts_with("ppt/media/") && !n.contains("handoutMaster") && !n.ends_with("theme2.xml")).collect();
    assert!(lost.is_empty(), "parts lost in the round trip: {lost:?}");
    assert!(!b.iter().any(|n| n.contains("handoutMaster")), "a part with a dangling relationship is not written");
    assert!(!pres_rels_has_handout(&out), "nor is the relationship to it");
    assert_eq!(b.iter().filter(|n| n.starts_with("ppt/media/")).count(), 1, "one copy of the preview picture");

    // Relationships point at the kept parts again.
    assert_eq!(rel_target(&part(&out, "_rels/.rels"), "rId4"), "docProps/thumbnail.jpeg");
    assert!(part(&out, "_rels/.rels").contains("relationships/metadata/thumbnail"), "the package-level type URI is kept whole");
    let pres_rels = part(&out, "ppt/_rels/presentation.xml.rels");
    assert!(pres_rels.contains(r#"Target="printerSettings/printerSettings1.bin""#), "{pres_rels}");
    let slide_xml = part(&out, "ppt/slides/slide1.xml");
    let slide_rels = part(&out, "ppt/slides/_rels/slide1.xml.rels");
    assert_eq!(
        rel_target(&slide_rels, attr(&slide_xml, "<p:oleObj", "r:id")),
        "../embeddings/oleObject1.bin",
        "the OLE frame is written as it was, under a new id"
    );
    assert_eq!(rel_target(&slide_rels, attr(&slide_xml, "<a:blip", "r:embed")), "../media/image1.png", "its preview points at the media item");
    let chart_xml = part(&out, "ppt/charts/chart1.xml");
    assert_eq!(
        rel_target(&part(&out, "ppt/charts/_rels/chart1.xml.rels"), attr(&chart_xml, "<c:externalData", "r:id")),
        "../embeddings/Microsoft_Excel_Sheet1.xlsx"
    );
    let ct = part(&out, "[Content_Types].xml");
    assert_eq!(content_type(&ct, "/ppt/embeddings/oleObject1.bin"), "application/vnd.openxmlformats-officedocument.oleObject");
    assert_eq!(
        content_type(&ct, "/ppt/printerSettings/printerSettings1.bin"),
        "application/vnd.openxmlformats-officedocument.presentationml.printerSettings"
    );
    assert_eq!(content_type(&ct, "/ppt/diagrams/data1.xml"), "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml");
    assert_eq!(content_type(&ct, "/docProps/thumbnail.jpeg"), "image/jpeg");

    // A second round trip is stable.
    let again = deckcraft_pptx::export(&deckcraft_pptx::import(&out).expect("re-import")).expect("re-export");
    assert_eq!(names(&again), b);
}

/// An edited chart is written from the model: the relationship to its original workbook goes
/// (the data differ now); the workbook bytes stay in the package as an orphan part.
#[test]
fn an_edited_chart_drops_the_relationship_to_its_stale_workbook() {
    let mut p = deckcraft_pptx::import(&deck()).expect("import");
    let slide = std::sync::Arc::make_mut(&mut p.slides[0]);
    for s in &mut slide.shapes {
        if let ShapeKind::Chart(c) = &mut s.kind {
            c.title = Some("Edited".into());
        }
    }
    let out = deckcraft_pptx::export(&p).expect("export");
    assert!(!part(&out, "ppt/charts/chart1.xml").contains("externalData"));
    assert!(!names(&out).contains("ppt/charts/_rels/chart1.xml.rels"));
}

/// Every deck of a corpus (`DECKCRAFT_PPTX_CORPUS=<dir>`) round-trips without losing a part;
/// the table of part counts is printed. Skipped when the variable is unset.
#[test]
fn corpus_round_trips_lose_no_part() {
    let Ok(dir) = std::env::var("DECKCRAFT_PPTX_CORPUS") else {
        eprintln!("DECKCRAFT_PPTX_CORPUS unset; corpus round trip skipped");
        return;
    };
    let mut files: Vec<_> = std::fs::read_dir(&dir)
        .expect("corpus dir")
        .filter_map(|e| e.ok().map(|e| e.path()))
        .filter(|p| p.extension().is_some_and(|e| e.eq_ignore_ascii_case("pptx")))
        .collect();
    files.sort();
    assert!(!files.is_empty(), "no .pptx in {dir}");
    let mut failures = vec![];
    println!("{:<44} {:>5} {:>5} {:>5}", "deck", "in", "out", "lost");
    for f in &files {
        let src = std::fs::read(f).expect("read");
        let out = deckcraft_pptx::export(&deckcraft_pptx::import(&src).expect("import")).expect("export");
        let (a, b) = (names(&src), names(&out));
        let lost: Vec<_> = a.difference(&b).cloned().collect();
        println!("{:<44} {:>5} {:>5} {:>5}", f.file_name().and_then(|n| n.to_str()).unwrap_or("?"), a.len(), b.len(), lost.len());
        if !lost.is_empty() {
            failures.push((f.clone(), lost));
        }
    }
    assert!(failures.is_empty(), "parts lost: {failures:?}");
}
