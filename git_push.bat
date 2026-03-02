@echo off
:: 先從雲端抓取更新
git pull origin main
git add .
set /p msg="請輸入提交訊息 (Commit Message): "
git commit -m "%msg%"
:: 上傳
git push origin main
echo.
echo 更新並上傳完成！
pause