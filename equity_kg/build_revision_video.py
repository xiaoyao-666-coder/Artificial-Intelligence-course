"""Build a narrated video from an inspected PowerPoint copy and matching WAVs."""
from pathlib import Path
import argparse
import hashlib
import json
import math
import re
import subprocess
import wave

from package_submission import ffmpeg_path, inspect_video

ROOT = Path(__file__).resolve().parent


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--qa-dir', type=Path, default=ROOT / 'qa/ppt-revision')
    parser.add_argument('--output', type=Path, default=ROOT / 'deliverables/PPT汇报视频_修订版.mp4')
    args = parser.parse_args()
    qa = args.qa_dir.resolve()
    output = args.output.resolve()
    content = json.loads((qa / 'content.json').read_text(encoding='utf-8-sig'))
    audit = json.loads((qa / 'powerpoint-audit.json').read_text(encoding='utf-8-sig'))
    audio_manifest = json.loads((qa / 'audio-manifest.json').read_text(encoding='utf-8-sig'))
    if digest(qa / 'content.json') != audio_manifest['content_sha256']:
        raise RuntimeError('Narration changed; regenerate the audio first.')
    if digest(qa / 'content.json') != audit['content_sha256'].lower():
        raise RuntimeError('Content changed since slide export; revise the deck first.')
    audio_hashes = {item['slide']: item['sha256'] for item in audio_manifest['audio']}
    deck = ROOT / content['output']
    if digest(deck) != audit['target_sha256'].lower():
        raise RuntimeError('Deck changed since slide export; export and inspect again.')
    slides = sorted((qa / 'slides').glob('*.PNG'), key=lambda p: int(re.search(r'\d+', p.stem).group()))
    if len(slides) != len(content['slides']):
        raise RuntimeError('Slide count differs from narration.')
    fps, speed, pause = 12, 1.08, 0.8
    segments = []
    for entry, slide in zip(content['slides'], slides):
        number = entry['slide']
        if int(re.search(r'\d+', slide.stem).group()) != number:
            raise RuntimeError('Slide order differs from narration.')
        audio = qa / 'audio' / f'voice-{number}.wav'
        if digest(audio) != audio_hashes.get(number):
            raise RuntimeError(f'Slide {number} audio changed since speech generation.')
        with wave.open(str(audio)) as wav:
            raw_seconds = wav.getnframes() / wav.getframerate()
        frames = math.ceil((raw_seconds / speed + pause) * fps)
        segments.append({'slide': number, 'seconds': frames / fps,
                         'image': slide, 'audio': audio, 'part': qa / f'talk-{number}.mp4'})
    if sum(s['seconds'] for s in segments) >= 180:
        raise RuntimeError('Narration exceeds three minutes; shorten it before encoding.')
    ffmpeg = ffmpeg_path()

    def run(argv):
        subprocess.run([ffmpeg, '-hide_banner', '-loglevel', 'error', '-y', *argv],
                       check=True, capture_output=True, timeout=240)

    for s in segments:
        run(['-loop', '1', '-framerate', str(fps), '-i', str(s['image']), '-i', str(s['audio']),
             '-t', str(s['seconds']), '-vf', 'scale=1600:900,format=yuv420p',
             '-af', f'atempo={speed},apad', '-r', str(fps), '-c:v', 'libx264', '-threads', '2',
             '-preset', 'fast', '-crf', '21', '-c:a', 'aac', '-ar', '48000', '-ac', '2',
             '-b:a', '160k', '-movflags', '+faststart', str(s['part'])])
        print(f"Encoded slide {s['slide']}: {s['seconds']:.2f}s", flush=True)
    listing = qa / 'revision-concat.txt'
    listing.write_text('\n'.join(f"file '{s['part'].name}'" for s in segments), encoding='utf-8')
    temporary = qa / 'video-candidate.mp4'
    run(['-f', 'concat', '-safe', '0', '-i', str(listing), '-c', 'copy', '-movflags', '+faststart', str(temporary)])
    info = inspect_video(temporary)
    if not (0 < info['seconds'] < 180 and info['audio']):
        raise RuntimeError('Final video must have audio and be under three minutes.')
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary.replace(output)
    start = 0.0
    evidence = []
    for s in segments:
        evidence.append({'slide': s['slide'], 'start': start, 'seconds': s['seconds'],
                         'image_sha256': digest(s['image']), 'audio_sha256': digest(s['audio'])})
        start += s['seconds']
    manifest = {'deck_sha256': digest(deck), 'video_sha256': digest(output),
                'content_sha256': digest(qa / 'content.json'), 'video': info,
                'speed': speed, 'pause_seconds': pause, 'segments': evidence}
    (qa / 'video-manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(info, ensure_ascii=False), flush=True)
    print(output, flush=True)


if __name__ == '__main__':
    main()
