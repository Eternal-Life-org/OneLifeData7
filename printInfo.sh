#!/bin/bash

# ==============================================================================
# 聚合打印脚本 - 游戏数据查看工具
# 用法：./print_menu.sh
# ==============================================================================

# 检查 bc 是否安装（部分脚本需要）
if ! command -v bc &> /dev/null; then
    echo "警告：未找到 'bc' 命令，部分功能可能无法正常工作。请安装 bc（例如：apt install bc）。"
fi

# ------------------------------------------------------------------------------
# 1. printAutoDecays
# ------------------------------------------------------------------------------
printAutoDecays() {
    for f in transitions/*; do
        if [[ $f == transitions/*.txt ]]; then
            if [ -e "$f" ]; then
                actorID=$(echo "$f" | sed 's/.*\///' | sed 's/_.*//')
                targetID=$(echo "$f" | sed 's/.*\///' | sed 's/\..*//' | sed 's/[^_]*_//')
                newActorID=$(cat "$f" | sed 's/\s.*//')
                newTargetID=$(cat "$f" | sed 's/[^ ]* //' | sed 's/\s.*//')
                decayTime=$(cat "$f" | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/\s.*//')
                lastUse=0
                if [[ $targetID == *"L"* ]]; then
                    lastUse=1
                    targetID=$(echo "$targetID" | sed 's/_.*//')
                fi
                decayString=""
                if [[ $actorID == -1 ]]; then
                    decayString="($decayTime seconds)"
                    target=$(cat "objects/$targetID.txt" | sed -n 2p)
                    target="\"$target\""
                    newTarget=""
                    if [[ $newTargetID == 0 ]]; then
                        newTarget="[NOTHING]"
                    else
                        newTarget=$(cat "objects/$newTargetID.txt" | sed -n 2p)
                        newTarget="\"$newTarget\""
                    fi
                    echo "$target  =>  $newTarget  $decayString"
                fi
            fi
        fi
    done
}

# ------------------------------------------------------------------------------
# 2. printBlockingObjects
# ------------------------------------------------------------------------------
printBlockingObjects() {
    for f in `ls -v objects/*`; do
        if [[ $f == objects/*.txt ]] && ! [[ $f == objects/nextObjectNumber.txt ]]; then
            name=$(cat "$f" | sed -n 2p)
            blocking=$(cat "$f" | sed -n 7p)
            block=$(echo "$blocking" | grep -o -E '[01]' | head -1)
            if (( $(bc <<< "$block > 0") )); then
                echo -n -e "$name\n"
            fi
        fi
    done
}

# ------------------------------------------------------------------------------
# 3. printContainableObjects
# ------------------------------------------------------------------------------
printContainableObjects() {
    for f in `ls -v objects/*`; do
        if [[ $f == objects/*.txt ]] && ! [[ $f == objects/nextObjectNumber.txt ]]; then
            name=$(cat "$f" | sed -n 2p)
            containSize=$(cat "$f" | sed -n 4p)
            while read -r line; do
                if [[ $line == containable* ]]; then
                    cont=$(echo "$line" | grep -o -E '[01]' | head -1)
                    size=$(echo "$containSize" | grep -o -E '[0-9]' | head -1)
                    if (( $(bc <<< "$cont == 1") )); then
                        echo -e "$name\n  | size = $size"
                    fi
                    break
                fi
            done < "$f"
        fi
    done
}

# ------------------------------------------------------------------------------
# 4. printContainerObjects
# ------------------------------------------------------------------------------
printContainerObjects() {
    for f in `ls -v objects/*`; do
        if [[ $f == objects/*.txt ]] && ! [[ $f == objects/nextObjectNumber.txt ]]; then
            name=$(cat "$f" | sed -n 2p)
            containable=$(cat "$f" | sed -n 3p)
            containSize=$(cat "$f" | sed -n 4p)
            cont=$(echo "$containable" | grep -o -E '[01]' | head -1)
            containableSize=$(echo "$containSize" | grep -o -E '[0-9]' | head -1)
            numSlots=0
            slotSize=0
            while read -r line; do
                if [[ $line == numSlots* ]]; then
                    numSlots=$(echo "$line" | grep -o -E '[0-9]' | head -1)
                fi
                if [[ $line == slotSize* ]]; then
                    slotSize=$(echo "$line" | grep -o -E '[0-9]' | head -1)
                fi
            done < "$f"
            if (( $(bc <<< "$numSlots > 0") )); then
                echo -n -e "$name\n  | $numSlots slots,  size = $slotSize"
                if (( $(bc <<< "$containable > 0") )); then
                    echo -n -e "\n  | containableSize = $containableSize"
                fi
                echo ""
            fi
        fi
    done
}

# ------------------------------------------------------------------------------
# 5. printDeadlyObjects
# ------------------------------------------------------------------------------
printDeadlyObjects() {
    for f in `ls -v objects/*`; do
        if [[ $f == objects/*.txt ]] && ! [[ $f == objects/nextObjectNumber.txt ]]; then
            name=$(cat "$f" | sed -n 2p)
            while read -r line; do
                if [[ $line == deadlyDistance* ]]; then
                    dist=$(echo "$line" | grep -o -E '[0-9.]+' | head -1)
                    if (( $(bc <<< "$dist > 0") )); then
                        echo -e "$dist\t $name"
                    fi
                    break
                fi
            done < "$f"
        fi
    done
}

# ------------------------------------------------------------------------------
# 6. printHoldableMultiSpriteObjects
# ------------------------------------------------------------------------------
printHoldableMultiSpriteObjects() {
    for f in `ls -v objects/*`; do
        if [[ $f == objects/*.txt ]] && ! [[ $f == objects/nextObjectNumber.txt ]]; then
            name=$(cat "$f" | sed -n 2p)
            while read -r line; do
                isPerm=0
                if [[ $line == permanent* ]]; then
                    perm=$(echo "$line" | grep -o -E '[01]' | head -1)
                    if (( $(bc <<< "$perm == 1") )); then
                        break
                    fi
                fi
                if [[ $line == numSprites* ]]; then
                    numSprites=$(echo "$line" | grep -o -E '[0-9]+' | head -1)
                    if (( $(bc <<< "$numSprites > 1") )); then
                        echo -e "($numSprites) \t $name"
                    fi
                    break
                fi
            done < "$f"
        fi
    done
}

# ------------------------------------------------------------------------------
# 7. printHoldableObjects
# ------------------------------------------------------------------------------
printHoldableObjects() {
    for f in `ls -v objects/*`; do
        if [[ $f == objects/*.txt ]] && ! [[ $f == objects/nextObjectNumber.txt ]]; then
            name=$(cat "$f" | sed -n 2p)
            while read -r line; do
                if [[ $line == permanent* ]]; then
                    perm=$(echo "$line" | grep -o -E '[01]' | head -1)
                    if (( $(bc <<< "$perm == 0") )); then
                        echo -e "$name"
                    fi
                    break
                fi
            done < "$f"
        fi
    done
}

# ------------------------------------------------------------------------------
# 8. printLastUseActor
# ------------------------------------------------------------------------------
printLastUseActor() {
    for f in transitions/*; do
        if [[ $f == transitions/*LA*.txt ]]; then
            if [ -e "$f" ]; then
                actorID=$(echo "$f" | sed 's/.*\///' | sed 's/_.*//')
                targetID=$(echo "$f" | sed 's/.*\///' | sed 's/\..*//' | sed 's/[^_]*_//')
                newActorID=$(cat "$f" | sed 's/\s.*//')
                newTargetID=$(cat "$f" | sed 's/[^ ]* //' | sed 's/\s.*//')
                decayTime=$(cat "$f" | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/\s.*//')
                reverseActor=$(cat "$f" | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/\s.*//')
                reverseTarget=$(cat "$f" | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/\s.*//')
                lastUseActor=0
                lastUseTarget=0
                if [[ $targetID == *"LA"* ]]; then
                    lastUseActor=1
                fi
                if [[ $targetID == *"LT"* ]]; then
                    lastUseTarget=1
                fi
                if [[ $targetID == *"L."* ]]; then
                    lastUseTarget=1
                fi
                if [[ $lastUseActor == 1 ]] || [[ $lastUseTarget == 1 ]]; then
                    targetID=$(echo "$targetID" | sed 's/_.*//')
                fi
                decayString=""
                actor=""
                if [[ $actorID == -1 ]]; then
                    actor="[DECAY]"
                    decayString="($decayTime seconds)"
                elif [[ $actorID == 0 ]]; then
                    actor="[HAND]"
                elif [[ $actorID == -2 ]]; then
                    actor="[DEFAULT]"
                else
                    actor=$(cat "objects/$actorID.txt" | sed -n 2p)
                    actor="\"$actor\""
                fi
                target=""
                if [[ $targetID == -1 ]] && [[ $newTargetID == 0 ]]; then
                    target="[USE/EAT]"
                elif [[ $targetID == -1 ]] && [[ $newTargetID != 0 ]]; then
                    target="[BARE-GROUND]"
                elif [[ $targetID == 0 ]]; then
                    target="[ON-PERSON]"
                else
                    target=$(cat "objects/$targetID.txt" | sed -n 2p)
                    target="\"$target\""
                fi
                newActor=""
                if [[ $newActorID == 0 ]]; then
                    newActor="[NOTHING]"
                else
                    newActor=$(cat "objects/$newActorID.txt" | sed -n 2p)
                    newActor="\"$newActor\""
                fi
                newTarget=""
                if [[ $newTargetID == 0 ]]; then
                    newTarget="[NOTHING]"
                else
                    newTarget=$(cat "objects/$newTargetID.txt" | sed -n 2p)
                    newTarget="\"$newTarget\""
                fi
                lastUseString=""
                if [[ $lastUseActor == 1 ]]; then
                    lastUseString="$lastUseString(Last Use Actor) "
                fi
                if [[ $lastUseTarget == 1 ]]; then
                    lastUseString="$lastUseString(Last Use Target) "
                fi
                reverseString=""
                if [[ $reverseActor == 1 ]]; then
                    reverseString="$reverseString(reverse actor) "
                fi
                if [[ $reverseTarget == 1 ]]; then
                    reverseString="$reverseString(reverse target) "
                fi
                echo "  $lastUseString$actor  +  $target   =   $newActor  +  $newTarget  $decayString   $reverseString"
            fi
        fi
    done
}

# ------------------------------------------------------------------------------
# 9. printLastUseTarget
# ------------------------------------------------------------------------------
printLastUseTarget() {
    for f in transitions/*; do
        if [[ $f == transitions/*LT*.txt ]]; then
            if [ -e "$f" ]; then
                actorID=$(echo "$f" | sed 's/.*\///' | sed 's/_.*//')
                targetID=$(echo "$f" | sed 's/.*\///' | sed 's/\..*//' | sed 's/[^_]*_//')
                newActorID=$(cat "$f" | sed 's/\s.*//')
                newTargetID=$(cat "$f" | sed 's/[^ ]* //' | sed 's/\s.*//')
                decayTime=$(cat "$f" | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/\s.*//')
                reverseActor=$(cat "$f" | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/\s.*//')
                reverseTarget=$(cat "$f" | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/\s.*//')
                lastUseActor=0
                lastUseTarget=0
                if [[ $targetID == *"LA"* ]]; then
                    lastUseActor=1
                fi
                if [[ $targetID == *"LT"* ]]; then
                    lastUseTarget=1
                fi
                if [[ $targetID == *"L."* ]]; then
                    lastUseTarget=1
                fi
                if [[ $lastUseActor == 1 ]] || [[ $lastUseTarget == 1 ]]; then
                    targetID=$(echo "$targetID" | sed 's/_.*//')
                fi
                decayString=""
                actor=""
                if [[ $actorID == -1 ]]; then
                    actor="[DECAY]"
                    decayString="($decayTime seconds)"
                elif [[ $actorID == 0 ]]; then
                    actor="[HAND]"
                elif [[ $actorID == -2 ]]; then
                    actor="[DEFAULT]"
                else
                    actor=$(cat "objects/$actorID.txt" | sed -n 2p)
                    actor="\"$actor\""
                fi
                target=""
                if [[ $targetID == -1 ]] && [[ $newTargetID == 0 ]]; then
                    target="[USE/EAT]"
                elif [[ $targetID == -1 ]] && [[ $newTargetID != 0 ]]; then
                    target="[BARE-GROUND]"
                elif [[ $targetID == 0 ]]; then
                    target="[ON-PERSON]"
                else
                    target=$(cat "objects/$targetID.txt" | sed -n 2p)
                    target="\"$target\""
                fi
                newActor=""
                if [[ $newActorID == 0 ]]; then
                    newActor="[NOTHING]"
                else
                    newActor=$(cat "objects/$newActorID.txt" | sed -n 2p)
                    newActor="\"$newActor\""
                fi
                newTarget=""
                if [[ $newTargetID == 0 ]]; then
                    newTarget="[NOTHING]"
                else
                    newTarget=$(cat "objects/$newTargetID.txt" | sed -n 2p)
                    newTarget="\"$newTarget\""
                fi
                lastUseString=""
                if [[ $lastUseActor == 1 ]]; then
                    lastUseString="$lastUseString(Last Use Actor) "
                fi
                if [[ $lastUseTarget == 1 ]]; then
                    lastUseString="$lastUseString(Last Use Target) "
                fi
                reverseString=""
                if [[ $reverseActor == 1 ]]; then
                    reverseString="$reverseString(reverse actor) "
                fi
                if [[ $reverseTarget == 1 ]]; then
                    reverseString="$reverseString(reverse target) "
                fi
                echo "  $lastUseString$actor  +  $target   =   $newActor  +  $newTarget  $decayString   $reverseString"
            fi
        fi
    done
}

# ------------------------------------------------------------------------------
# 10. printNaturalObjects
# ------------------------------------------------------------------------------
printNaturalObjects() {
    for f in `ls -v objects/*`; do
        if [[ $f == objects/*.txt ]] && ! [[ $f == objects/nextObjectNumber.txt ]]; then
            name=$(cat "$f" | sed -n 2p)
            while read -r line; do
                if [[ $line == mapChance* ]]; then
                    chance=$(echo "$line" | grep -o -E '[0-9.]+' | head -1)
                    biomes=$(echo "$line" | grep -o -E 'biomes_[0-9,]+' | tail -1)
                    if (( $(bc <<< "$chance > 0") )); then
                        echo -e "$chance  $biomes \t $name"
                    fi
                    break
                fi
            done < "$f"
        fi
    done
}

# ------------------------------------------------------------------------------
# 11. printNonContaiableObjects (原拼写)
# ------------------------------------------------------------------------------
printNonContaiableObjects() {
    for f in `ls -v objects/*`; do
        if [[ $f == objects/*.txt ]] && ! [[ $f == objects/nextObjectNumber.txt ]]; then
            name=$(cat "$f" | sed -n 2p)
            containSize=$(cat "$f" | sed -n 4p)
            permanent=$(cat "$f" | sed -n 5p)
            while read -r line; do
                if [[ $line == containable* ]]; then
                    cont=$(echo "$line" | grep -o -E '[01]' | head -1)
                    size=$(echo "$containSize" | grep -o -E '[0-9]' | head -1)
                    perm=$(echo "$permanent" | grep -o -E '[01]' | head -1)
                    if (( $(bc <<< "$cont == 0") )) && (( $(bc <<< "$perm == 0") )) && ! [[ $name == *Female* ]]  && ! [[ $name == *Male* ]] && ! [[ $name == *@* ]]; then
                        echo -e "$name\n"
                    fi
                    break
                fi
            done < "$f"
        fi
    done
}

# ------------------------------------------------------------------------------
# 12. printUseableObjects
# ------------------------------------------------------------------------------
printUseableObjects() {
    for f in `ls -v objects/*`; do
        if [[ $f == objects/*.txt ]] && ! [[ $f == objects/nextObjectNumber.txt ]]; then
            name=$(cat "$f" | sed -n 2p)
            while read -r line; do
                if [[ $line == numUses* ]]; then
                    uses=$(echo "$line" | grep -o -E '[0-9]+' | head -1)
                    useChance=$(echo "$line" | sed 's/.*,//' | head -1)
                    if [[ $useChance == numUses* ]]; then
                        useChance="1.0"
                    fi
                    if (( $(bc <<< "$uses > 1") )); then
                        echo -e "$name\n  | uses = $uses   [$useChance]"
                    fi
                    break
                fi
            done < "$f"
        fi
    done
}

# ------------------------------------------------------------------------------
# 13. printReverseUseDecayTrans
# ------------------------------------------------------------------------------
printReverseUseDecayTrans() {
    echo
    echo "Reverse use decay trans:"
    ls -1 transitions | while read x; do
        f="transitions/$x"
        if [[ $f == transitions/*.txt ]]; then
            if [ -e "$f" ]; then
                actorID=$(echo "$f" | sed 's/.*\///' | sed 's/_.*//')
                targetID=$(echo "$f" | sed 's/.*\///' | sed 's/\..*//' | sed 's/[^_]*_//')
                newActorID=$(cat "$f" | sed 's/\s.*//')
                newTargetID=$(cat "$f" | sed 's/[^ ]* //' | sed 's/\s.*//')
                decayTime=$(cat "$f" | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/\s.*//')
                reverseUseActor=$(cat "$f" | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/\s.*//')
                reverseUseTarget=$(cat "$f" | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/\s.*//')
                noUseActor=$(cat "$f" | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/\s.*//')
                noUseTarget=$(cat "$f" | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/[^ ]* //' | sed 's/\s.*//')
                lastUseActor=0
                lastUseTarget=0
                if [[ $targetID == *"LA"* ]]; then
                    lastUseActor=1
                elif [[ $targetID == *"LT"* ]]; then
                    lastUseTarget=1
                elif [[ $targetID == *"L"* ]]; then
                    lastUseTarget=1
                fi
                if [[ $lastUseActor == 1 ]] || [[ $lastUseTarget == 1 ]]; then
                    targetID=$(echo "$targetID" | sed 's/_.*//')
                fi
                decayString=""
                actor=""
                if [[ $actorID == -1 ]]; then
                    actor="[DECAY]"
                    decayString="($decayTime seconds)"
                elif [[ $actorID == 0 ]]; then
                    actor="[HAND]"
                elif [[ $actorID == -2 ]]; then
                    actor="[DEFAULT]"
                else
                    actor=$(cat "objects/$actorID.txt" | sed -n 2p)
                    actor="\"$actor\""
                fi
                target=""
                if [[ $targetID == -1 ]] && [[ $newTargetID == 0 ]]; then
                    target="[USE/EAT]"
                elif [[ $targetID == -1 ]] && [[ $newTargetID != 0 ]]; then
                    target="[BARE-GROUND]"
                elif [[ $targetID == 0 ]]; then
                    target="[ON-PERSON]"
                else
                    target=$(cat "objects/$targetID.txt" | sed -n 2p)
                    target="\"$target\""
                fi
                newActor=""
                if [[ $newActorID == 0 ]]; then
                    newActor="[NOTHING]"
                else
                    newActor=$(cat "objects/$newActorID.txt" | sed -n 2p)
                    newActor="\"$newActor\""
                fi
                newTarget=""
                if [[ $newTargetID == 0 ]]; then
                    newTarget="[NOTHING]"
                else
                    newTarget=$(cat "objects/$newTargetID.txt" | sed -n 2p)
                    newTarget="\"$newTarget\""
                fi
                lastUseString=""
                if [[ $lastUseActor == 1 ]]; then
                    lastUseString="$lastUseString(Last Use Actor) "
                fi
                if [[ $lastUseTarget == 1 ]]; then
                    lastUseString="$lastUseString(Last Use Target) "
                fi
                reverseUseString=""
                if [[ $reverseUseActor == 1 ]]; then
                    reverseUseString="$reverseUseString(Reverse Use Actor) "
                fi
                if [[ $reverseUseTarget == 1 ]]; then
                    reverseUseString="$reverseUseString(Reverse Use Target) "
                fi
                noUseString=""
                if [[ $noUseActor == 1 ]]; then
                    noUseString="$noUseString(No Use Actor) "
                fi
                if [[ $noUseTarget == 1 ]]; then
                    noUseString="$noUseString(No Use Target) "
                fi
                if [[ $actorID == -1 ]]; then
                    if [[ $reverseUseActor == 1 ]] || [[ $reverseUseTarget == 1 ]]; then
                        echo "  $lastUseString$actor  +  $target   =   $newActor  +  $newTarget  $decayString  $reverseUseString $noUseString"
                        echo
                        echo
                    fi
                fi
            else
                echo "$f removed"
            fi
        fi
    done
}

# ------------------------------------------------------------------------------
# 菜单主程序
# ------------------------------------------------------------------------------
show_menu() {
    clear
    echo "============================================"
    echo "         游戏数据打印工具 - 菜单           "
    echo "============================================"
    echo "  1) 自动衰减 (AutoDecays)"
    echo "  2) 阻塞物体 (BlockingObjects)"
    echo "  3) 可收纳物体 (ContainableObjects)"
    echo "  4) 容器物体 (ContainerObjects)"
    echo "  5) 致命物体 (DeadlyObjects)"
    echo "  6) 多精灵可持物体 (HoldableMultiSpriteObjects)"
    echo "  7) 可持物体 (HoldableObjects)"
    echo "  8) 最后使用 - 角色 (LastUseActor)"
    echo "  9) 最后使用 - 目标 (LastUseTarget)"
    echo " 10) 自然生成物体 (NaturalObjects)"
    echo " 11) 不可收纳物体 (NonContainableObjects)"
    echo " 12) 可用物体 (UseableObjects)"
    echo " 13) 反向使用衰减转换 (ReverseUseDecayTrans)"
    echo "  0) 退出"
    echo "============================================"
    echo -n "请输入选项 [0-13]: "
}

while true; do
    show_menu
    read choice
    case $choice in
        1) printAutoDecays ;;
        2) printBlockingObjects ;;
        3) printContainableObjects ;;
        4) printContainerObjects ;;
        5) printDeadlyObjects ;;
        6) printHoldableMultiSpriteObjects ;;
        7) printHoldableObjects ;;
        8) printLastUseActor ;;
        9) printLastUseTarget ;;
        10) printNaturalObjects ;;
        11) printNonContaiableObjects ;;
        12) printUseableObjects ;;
        13) printReverseUseDecayTrans ;;
        0) echo "退出程序。"; exit 0 ;;
        *) echo "无效选项，请重新输入。"; sleep 1 ;;
    esac
    echo
    echo -n "按 Enter 键返回菜单..."
    read
done
