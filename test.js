const { createReadStream } = require('fs');

async function StreamReadAsync(stream) {

    const buffers = [];

    for await (const chunk of stream) {
        buffers.push(chunk);
    }

    return Buffer.concat(buffers);

}

StreamReadAsync(createReadStream('./2017.xlsx',{
    autoClose:true
}))
.then(result=>console.log(result.length))
.catch(error=>console.log(error));