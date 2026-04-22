"""Generate offline Japanese TTS audio and embed into pptx."""
from __future__ import annotations

import subprocess
from pathlib import Path

from pptx import Presentation
from pptx.util import Cm

JTALK_DICT = "/var/lib/mecab/dic/open-jtalk/naist-jdic"
JTALK_VOICE = "/usr/share/hts-voice/nitech-jp-atr503-m001/nitech_jp_atr503_m001.htsvoice"


def tts(text: str, out_path: Path, *, speed: float = 1.0,
        pitch: float = 0.0, volume: float = 0.0) -> None:
    """Render text to wav via open_jtalk."""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wav = out_path.with_suffix(".wav")
    proc = subprocess.run(
        [
            "open_jtalk",
            "-x", JTALK_DICT,
            "-m", JTALK_VOICE,
            "-ow", str(wav),
            "-r", str(speed),
            "-fm", str(pitch),
            "-g", str(volume),
        ],
        input=text.encode("utf-8"),
        capture_output=True,
        check=True,
    )
    # Convert wav -> m4a (smaller, broadly compatible with pptx)
    if out_path.suffix.lower() in {".m4a", ".mp3"}:
        codec = "aac" if out_path.suffix.lower() == ".m4a" else "libmp3lame"
        subprocess.run(
            [
                "ffmpeg", "-y", "-loglevel", "error",
                "-i", str(wav),
                "-c:a", codec, "-b:a", "64k",
                str(out_path),
            ],
            check=True,
        )
        wav.unlink()


def set_note(slide, text: str) -> None:
    nt = slide.notes_slide.notes_text_frame
    nt.text = text


def add_audio(slide, audio_path: Path, *, position_cm=(0.3, 0.3)) -> None:
    """Embed audio as auto-play so it fires on slide enter."""
    from pptx.util import Cm as _Cm
    movie = slide.shapes.add_movie(
        str(audio_path),
        _Cm(position_cm[0]), _Cm(position_cm[1]),
        _Cm(1.0), _Cm(1.0),
        mime_type="audio/mp4" if audio_path.suffix.lower() == ".m4a" else "audio/mpeg",
    )
    _force_autoplay(slide, movie)


def _force_autoplay(slide, movie_shape) -> None:
    """Rewrite timing XML so the embedded media auto-starts on slide show.

    python-pptx's add_movie injects a default click-activated timing tree and
    marks the nvPr as a videoFile. We convert it to audioFile and swap the
    timing for an auto-play sequence targeted at this shape's spid.
    """
    from pptx.oxml.ns import qn
    from lxml import etree

    sp = movie_shape._element  # p:pic

    nvSpPr = sp.find(qn("p:nvPicPr"))
    cNvPr = nvSpPr.find(qn("p:cNvPr"))
    spid = cNvPr.get("id")

    nvPr = nvSpPr.find(qn("p:nvPr"))
    # python-pptx marks the media as a:videoFile; flip it to a:audioFile
    # so PowerPoint treats the shape as audio.
    vid = nvPr.find(qn("a:videoFile"))
    if vid is not None:
        rid = vid.get(qn("r:link")) or vid.get(qn("r:embed"))
        nvPr.remove(vid)
        audio_el = etree.SubElement(nvPr, qn("a:audioFile"))
        if rid:
            audio_el.set(qn("r:link"), rid)

    sld = slide._element  # p:sld
    for existing in sld.findall(qn("p:timing")):
        sld.remove(existing)

    timing_xml = f"""
    <p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:tnLst>
        <p:par>
          <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
            <p:childTnLst>
              <p:seq concurrent="1" nextAc="seek">
                <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                  <p:childTnLst>
                    <p:par>
                      <p:cTn id="3" fill="hold">
                        <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
                        <p:childTnLst>
                          <p:par>
                            <p:cTn id="4" fill="hold">
                              <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                              <p:childTnLst>
                                <p:par>
                                  <p:cTn id="5" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect">
                                    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                    <p:childTnLst>
                                      <p:cmd type="call" cmd="playFrom(0.0)">
                                        <p:cBhvr>
                                          <p:cTn id="6" dur="indefinite" fill="hold"/>
                                          <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
                                        </p:cBhvr>
                                      </p:cmd>
                                    </p:childTnLst>
                                  </p:cTn>
                                </p:par>
                              </p:childTnLst>
                            </p:cTn>
                          </p:par>
                        </p:childTnLst>
                      </p:cTn>
                    </p:par>
                  </p:childTnLst>
                </p:cTn>
                <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
                <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
              </p:seq>
              <p:audio><p:cMediaNode vol="80000"><p:cTn id="7" fill="hold" display="0"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cMediaNode></p:audio>
            </p:childTnLst>
          </p:cTn>
        </p:par>
      </p:tnLst>
    </p:timing>
    """.strip()
    sld.append(etree.fromstring(timing_xml))


def attach_narration(prs: Presentation, narrations: list[str],
                     audio_dir: Path, *, deck_key: str,
                     embed_audio: bool = True) -> None:
    """Add speaker notes and (optionally) embedded audio for each slide."""
    audio_dir.mkdir(parents=True, exist_ok=True)
    for i, slide in enumerate(prs.slides):
        script = narrations[i] if i < len(narrations) else ""
        if not script.strip():
            continue
        set_note(slide, script)
        if not embed_audio:
            continue
        audio_path = audio_dir / f"{deck_key}_{i+1:02d}.m4a"
        tts(script, audio_path)
        add_audio(slide, audio_path, position_cm=(32.3, 17.9))
