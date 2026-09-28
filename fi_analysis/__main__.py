"""python -m fi_analysis DECK.pptx -o results   (or --images DIR --map map.csv)"""
import argparse

from .pipeline import parse_roi, run
from .report import write_all


def main(argv=None):
    ap = argparse.ArgumentParser(prog="python -m fi_analysis", description=__doc__)
    ap.add_argument("deck", nargs="?", help="the .pptx with concentration labels above the images")
    ap.add_argument("--images", help="folder of raw images (use instead of a deck)")
    ap.add_argument("--map", help="CSV with columns file,conc[,unit] for --images")
    ap.add_argument("-o", "--out", default="results", help="output folder (default: results)")
    ap.add_argument("--roi", default="0.1,0.1,0.8,0.8",
                    help="x,y,w,h as fractions of each image; the same ROI is used for every image")
    ap.add_argument("--channel", default="auto", choices=["auto", "R", "G", "B", "gray"])
    ap.add_argument("--reference", help="CSV conc,fi[,sd] of the existing analysis (overrides slides)")
    ap.add_argument("--ref-slides", default="2,3", help="slides holding the existing analysis (default 2,3)")
    ap.add_argument("--model", default="auto", choices=["auto", "langmuir", "hill"])
    ap.add_argument("--tiles", type=int, default=3, help="ROI is split into tiles x tiles for spatial SD")
    ap.add_argument("--no-render", action="store_true", help="skip LibreOffice slide rendering")
    args = ap.parse_args(argv)
    if bool(args.deck) == bool(args.images) or (args.images and not args.map):
        ap.error("give either a deck, or --images together with --map")
    res = run(deck_path=args.deck, images_dir=args.images, map_csv=args.map, roi=parse_roi(args.roi),
              channel=args.channel, reference_csv=args.reference,
              ref_slides=tuple(int(s) for s in args.ref_slides.split(",") if s.strip()),
              model=args.model, tiles=args.tiles)
    write_all(res, args.out, args.deck or args.images, render=not args.no_render)
    print(f"done -> {args.out}/summary.md")


if __name__ == "__main__":
    main()
