# Build container

```
docker build -t homi/librsvg .
```

# Example use

```
docker run --rm -v "${PWD}:/data" -w /data homi/librsvg --background-color=white --width=4000px -f png -o output.png page7.svg
```