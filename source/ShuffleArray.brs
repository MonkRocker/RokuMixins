Function ShuffleArray(array as Object)
    'Function which performs a Fisher-Yates shuffle of a supplied array
    for i% = 0 to array.count() - 2 step 1
      j% = Rnd(array.count())
      'swap the two items
      tempItem = array[i%]
      array[i%] = array[j%]
      array[j%] = tempItem
    end for

    return array
End Function
