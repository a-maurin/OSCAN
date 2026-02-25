@echo off
REM Script pour recréer la branche main avec UN SEUL commit (code uniquement).
REM À lancer depuis la racine du depot OSCAN.

echo === Nettoyage de l'historique Git (code uniquement) ===

git checkout --orphan code-only
if errorlevel 1 ( echo ERREUR checkout orphan & pause & exit /b 1 )

git reset
git add .
git status
echo.
echo Verifiez ci-dessus : il ne doit PAS y avoir db_test, resultats, __pycache__
echo Si tout est OK, le commit va etre cree.
pause

git commit -m "Initial commit (code only, sans donnees confidentielles)"
if errorlevel 1 ( echo ERREUR commit & pause & exit /b 1 )

git branch -D main
git branch -m code-only main

echo.
echo Branche main recreee avec 1 seul commit.
echo Pour pousser vers GitHub : git push -u origin main --force
pause
