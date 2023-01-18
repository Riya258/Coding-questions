/*write a program to find the largest element from an array.*/
function riya()
 {
  const array=[2,5,8,9,3,6];
    my=array.sort();
    console.log("Sorted array is:"+my);
    let largest=0;
    for(let i=1;i<array.length;i++)
    {
        if(array[i]>largest)
            largest=array[i];
    }
    console.log("Largest element from this array is:"+largest);
 }
 return riya();

