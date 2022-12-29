# Google-Sheets-Scripts
<br>
Several different scripts written for optimizing various use cases in Googe Sheets

<br><br><br><br>

## Important Note!
Many of these scripts use the build-in `onEdit()` function. Google only will run 1 of these said functions in an intire project. So if you want to use multiple of these scripts, concider renaming the current `onEdit()` functions and using only 1 `onEdit()` function to run each of them. Don't forget to pass the parameter :)

```javascript
function MakeRoomFor_onEdit(e) {
  . . .
}
function TrackAndMatch_onEdit(e) {
  . . .
}

function onEdit(e) {
  MakeRoomFor_onEdit(e);
  TrackAndMatch_onEdit(e);
}
```
