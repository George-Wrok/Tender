@echo off
chcp 65001 > nul

echo ========================================
echo 選擇上傳帳號：
echo [1] George-Wrok
echo [2] shirleyhilti22348814
echo ========================================
set /p account_choice="輸入代號 (1 或 2): "

if "%account_choice%"=="2" (
    echo [切換為 shirleyhilti22348814]
    git config user.name "shirleyhilti22348814"
    git config user.email "273599169+shirleyhilti22348814@users.noreply.github.com"
    git remote set-url origin https://shirleyhilti22348814@github.com/George-Wrok/Tender.git
) else (
    echo [切換為 George-Wrok]
    git config user.name "George-Wrok"
    git config user.email "264807132+George-Wrok@users.noreply.github.com"
    git remote set-url origin https://George-Wrok@github.com/George-Wrok/Tender.git
)
echo ========================================

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