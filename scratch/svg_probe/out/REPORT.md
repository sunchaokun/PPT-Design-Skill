# SVG → editable-pptx probe report
Total compile (all cases): 68 ms

## pyramid
- compile: 13 ms | 6 shapes
- editable: True | pictures=False p:sp=6 groups=False
- features: gradient, polygon, text
- pixel diff(mean abs)=1.9  ink-IoU=0.98

## venn_evenodd
- compile: 9 ms | 4 shapes
- editable: True | pictures=False p:sp=4 groups=False
- features: path, text
- pixel diff(mean abs)=5.2  ink-IoU=0.83

## funnel
- compile: 9 ms | 8 shapes
- editable: True | pictures=False p:sp=8 groups=False
- features: gradient, polygon, text
- pixel diff(mean abs)=1.4  ink-IoU=0.98

## growth_curve
- compile: 13 ms | 5 shapes
- editable: True | pictures=False p:sp=6 groups=False
- features: circle, clipPath, gradient, line, path, text
- pixel diff(mean abs)=2.6  ink-IoU=0.91

## matrix_bcg
- compile: 21 ms | 13 shapes
- editable: True | pictures=False p:sp=13 groups=False
- features: circle, line, rect, text
- pixel diff(mean abs)=0.9  ink-IoU=0.91

## unsupported
- compile: 3 ms | 2 shapes
- editable: True | pictures=False p:sp=2 groups=False
- features: circle, image, rect
- pixel diff(mean abs)=8.0  ink-IoU=0.78
- boundaries: ['image element: UNSUPPORTED (would degrade to picture, refusing)', 'rect has filter=<url(#blur)>: UNSUPPORTED (refusing, no silent degrade)', 'circle has mask=<url(#m)>: UNSUPPORTED (refusing, no silent degrade)']
