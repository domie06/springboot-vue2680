const base = {
    get() {
        return {
            url : "http://localhost:8080/jiaoshixinxiguanli/",
            name: "jiaoshixinxiguanli",
            // 退出到首页链接
            indexUrl: 'http://localhost:8080/jiaoshixinxiguanli/front/index.html'
        };
    },
    getProjectName(){
        return {
            projectName: "教室信息管理系统"
        } 
    }
}
export default base
