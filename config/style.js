const defaultBorder = {
    left : {style : "thin"}, right : {style : "thin"}, top : {style : "thin"}, bottom : {style : "thin"}
}

module.exports = {
    nameHeader: {
        font: {
            bold: true
        },
        fill: {
            type : "pattern",
            patternType : "solid",
            fgColor: "#c6e0b4"
        },
        border : defaultBorder
    },
    nameHeaderValue: {
        font: {
            underline: true
        },
        fill: {
            type : "pattern",
            patternType : "solid",
            fgColor: "#c6e0b4"
        },
        border : defaultBorder
    },
    tableHeader: {
        alignment: {
            horizontal : "center"
        },
        font: {
            bold: true,
            color :"#ffffff"
        },
        fill: {
            type : "pattern",
            patternType : "solid",
            fgColor: "#c00000"
        },
        border : defaultBorder
    },
}