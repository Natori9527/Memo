#File reader
  lineNo=1
  while read -r line
  do
    echo "Line:$lineNo $line"
    ((lineNo = lineNo + 1))
  done < $TEMP_RESULT

#Trim
result='3!d' "  aa  "
${result//[[:blank:]]/}

#show one line in file
#delete other lines
sed '3!d' $TEMP_RESULT

#function return string
#use echo "aaaaa"
#return command only reutrn numeric as a result of process
function getAAA(){
  str="hello $1"
  echo $str
}

result=${getAAA "aaa"}
