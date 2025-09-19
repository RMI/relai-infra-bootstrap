import React from 'react';

/**
 * Renders information about the user obtained from MS Graph
 * @param props
 */
export const FilesData = (props) => {
    const data = props.graphData.value;
    const urls = data.map(d => d["@microsoft.graph.downloadUrl"]);

    async function get_content(url) {
        const config = {
            newlineDelimiter: " ",  // Separate new lines with a space instead of the default \n.
            ignoreNotes: true       // Ignore notes while parsing presentation files like pptx or odp.
        }
        const response = await fetch(url);
        const arrayBuffer = await response.arrayBuffer();
        const result = await officeParser.parseOfficeAsync(arrayBuffer, config);
        console.log(await result);
        return result;
    }

    window.urls = urls;
    console.log("window.urls = urls;");
    window.get_content = get_content;

    async function get_all_files_content(urls) {
        const result = await urls.map(async (url) => {
            let result;
            get_content(url)
                .then((response) => {
                    result = {
                        id: "e.id",
                        type: "file content",
                        date_time: "e.lastModifiedDateTime",
                        author: "e.from.user.displayName",
                        content: response || "",
                        subject: "e.subject"
                    };
                });
            console.log(await result);
            return await result;
        })
        return await result;
    }

    let file_content;
    get_all_files_content(urls).then((res) => file_content = res);

    console.log(file_content);
    window.file_content = file_content;

    return (
        <div id="files-div">
            <pre style={{textAlign: "left"}}>
                {JSON.stringify(file_content, null, 2) }
            </pre>
        </div>
    );
};
